using Microsoft.Win32;
using ReInvented.DataAccess.NameProviders;
using ReInvented.DataAccess.Services;
using ReInvented.Licensing.Core.Models;
using ReInvented.Licensing.Core.Services;
using ReInvented.Licensing.Core.ViewModels;
using ReInvented.Shared.Commands;
using ReInvented.Shared.Interfaces;
using ReInvented.Shared.Stores;
using System;
using System.Collections.Generic;
using System.IO;

using System.Windows;
using System.Windows.Input;

namespace XcelTools.Licensing.ViewModels
{
    public class RegistrationViewModel : ValidatablePropertyStore
    {
        #region Events

        public event Action Registered;
        public event Action Cancelled;

        #endregion

        #region Parameterized Constructor

        public RegistrationViewModel(IDialogService dialogService)
        {
            DialogService = dialogService ?? throw new ArgumentNullException(nameof(dialogService));
            Initialize();
        }

        #endregion

        #region Public Properties

        public string LicenseFilePath { get => Get<string>(); set { Set(value); RaisePropertyChanged(nameof(CanRegister)); } }

        public string RegistrationKey { get => Get<string>(); set { Set(value); RaisePropertyChanged(nameof(CanRegister)); } }

        public bool IsRegistered { get => Get<bool>(); set => Set(value); }

        public IDialogService DialogService { get => Get<IDialogService>(); private set => Set(value); }

        #endregion

        #region Commands

        public ICommand SelectLicenseFileCommand { get => Get<ICommand>(); set => Set(value); }

        public ICommand RegisterCommand { get => Get<ICommand>(); set => Set(value); }

        public ICommand GenerateLicenseInputCommand { get => Get<ICommand>(); set => Set(value); }

        public ICommand CancelCommand { get => Get<ICommand>(); set => Set(value); }

        #endregion

        #region Readonly Properties

        public bool CanRegister => !string.IsNullOrWhiteSpace(RegistrationKey) && !string.IsNullOrWhiteSpace(LicenseFilePath);

        #endregion

        #region Command Handlers

        private void OnSelectLicenseFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                Filter = $"License Files (*.{FileExtensions.XtractorLicense}) | *.{FileExtensions.XtractorLicense}",
                InitialDirectory = DirectoryPaths.UserLocalPublisher,
                Multiselect = false
            };
            if (openFileDialog.ShowDialog() == true)
            {
                LicenseFilePath = openFileDialog.FileName;
            }
        }

        private void OnRegister()
        {
            if (AuthenticationService.ValidateRegistrationKey(RegistrationKey, LicenseFilePath))
            {
                License license = PersistenceService.ReadFromFile<License>(LicenseFilePath);

                Registration registration;

                if (File.Exists(FilePaths.RegistrationFile))
                {
                    registration = PersistenceService.ReadFromFile<Registration>(FilePaths.RegistrationFile);
                }
                else
                {
                    registration = new Registration() { Records = new HashSet<RegistrationRecord>() };
                }

                var regRecord = new RegistrationRecord() { Module = license.ModuleType, LicenseFilePath = LicenseFilePath, RegistrationKey = RegistrationKey };

                registration.Records.Add(regRecord);

                if (LicenseFilePath != FilePaths.XtrLicenseFilePath)
                {
                    File.Copy(LicenseFilePath, FilePaths.XtrLicenseFilePath, true);
                }

                PersistenceService.WriteToFile(FilePaths.RegistrationFile, registration);

                MessageBox.Show($"The registration is successfully completed. You can start using the product. You can now start using {nameof(XcelTools)} AddIn.", "Registration successful", MessageBoxButton.OK);

                Registered?.Invoke();
            }
            else
            {
                MessageBox.Show("The registration key provided is invalid. Please contact the publisher.", "Invalid registration key", MessageBoxButton.OK);
            }
        }

        private void OnGenerateLicenseInput()
        {
            LicenseInputViewModel viewModel = new LicenseInputViewModel();
            _ = DialogService.ShowDialog(viewModel);
        }

        private void OnCancel()
        {
            Cancelled?.Invoke();
        }


        #endregion

        #region Abstract Methods Implementation

        protected void Initialize()
        {
            SelectLicenseFileCommand = new RelayCommand(OnSelectLicenseFile, true);
            RegisterCommand = new RelayCommand(OnRegister, true);
            GenerateLicenseInputCommand = new RelayCommand(OnGenerateLicenseInput, true);
            CancelCommand = new RelayCommand(OnCancel, true);
        }

        #endregion
    }
}
