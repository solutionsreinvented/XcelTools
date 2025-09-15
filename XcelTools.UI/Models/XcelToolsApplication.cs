using Microsoft.Office.Core;

using ReInvented.Licensing.Core.Enums;
using ReInvented.Licensing.Core.Models;
using ReInvented.Licensing.Core.Services;
using ReInvented.Licensing.Core.ViewModels;
using ReInvented.Licensing.Core.Views;
using ReInvented.Sections.Domain.Repositories;
using ReInvented.Shared.Dialogs;
using ReInvented.Shared.Interfaces;
using ReInvented.Shared.Services;
using ReInvented.Shared.ViewModels;

using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

using XcelTools.Licensing.Services;
using XcelTools.Licensing.ViewModels;
using XcelTools.Licensing.Views;
using XcelTools.UI.Views;

namespace XcelTools.UI.Models
{
    public class XcelToolsApplication
    {
        #region Constructors

        public XcelToolsApplication(COMAddIn addIn)
        {

            AddIn = addIn;
            ExcelApp = AddIn.Application as Microsoft.Office.Interop.Excel.Application;

            Application = new Application();

            Uri resourceUri = new Uri("/ReInvented.UI.Themes;component/Themes/ReInvented.UI.Themes.xaml", UriKind.Relative);
            ResourceDictionary resourceDict = new ResourceDictionary { Source = resourceUri };

            if (!Application.Resources.MergedDictionaries.Any(d => d.Source == resourceUri))
            {
                Application.Resources.MergedDictionaries.Add(resourceDict);
            }

            Dispatcher = Application.Dispatcher;
        }

        #endregion

        #region Private Properties

        private RegistrationView RegistrationView { get; set; }

        private IDialogService DialogService { get; set; }

        private System.Timers.Timer LicenseTimer { get; set; }

        private Application Application { get; set; }

        private Microsoft.Office.Interop.Excel.Application ExcelApp { get; set; }

        private Dispatcher Dispatcher { get; set; }

        private COMAddIn AddIn { get; set; }

        #endregion

        #region Main Method

        public void Start()
        {
            SplashScreenView splashScreen = new SplashScreenView();
            splashScreen.Show();
            LicenseTimer = new System.Timers.Timer(500);

            LicenseTimer.Elapsed += (sender, args) =>
            {
                LicenseTimer.Stop();

                Dispatcher dispatcher = Application.Current.Dispatcher;
                Dispatcher.Invoke(() => HandleUserAuthentication(splashScreen));
            };


            LicenseTimer.Start();
        }

        #endregion

        #region Private Helpers

        private void HandleUserAuthentication(SplashScreenView splashScreen)
        {
            DialogService = new DialogService();
            ConfigureDialogs();

            LicenseStatus licenseStatus = AuthenticationService.GetLicenseStatus(ApplicationModule.Xtractor, null);

            if (licenseStatus.IsValid)
            {
                Task.Run(() => SectionsRepository.Instance.GetSectionsLibrary());
            }
            else
            {
                DirectUserToRegistration();
            }

            splashScreen.Close();
        }

        private void DirectUserToRegistration()
        {
            Dispatcher.Invoke(() =>
            {
                RegistrationViewModel registerVM = new RegistrationViewModel(DialogService);
                RegistrationView = new RegistrationView { DataContext = registerVM };

                DialogService.SetOrChangeOwner(RegistrationView);

                registerVM.Cancelled -= OnRegistrationCancelled;
                registerVM.Cancelled += OnRegistrationCancelled;

                registerVM.Registered -= OnRegistrationCompleted;
                registerVM.Registered += OnRegistrationCompleted;

                //ExcelApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlMinimized;

                RegistrationView.Show();
            });
        }

        private void OnRegistrationCompleted()
        {
            //ExcelApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;

            RegistrationView.Close();
            Application.Shutdown();
        }

        private void ConfigureDialogs()
        {
            DialogService.Register<LicenseInputViewModel, LicenseInputGenerationView>();
            DialogService.Register<MessageViewModel, MessageDialog>();
        }

        #endregion

        #region Event Handlers

        private void OnRegistrationCancelled()
        {
            DeselectOrUnloadAddIn(AddIn);

            //ExcelApp.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal;
            Application.Shutdown();
        }

        private static void DeselectOrUnloadAddIn(COMAddIn addIn)
        {
            addIn.Connect = false;
            return;
        }

        #endregion
    }
}
