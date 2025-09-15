using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

using ReInvented.Shared.Services;

using System;
using System.Timers;

using XcelTools.Models;
using XcelTools.UI.Models;

namespace XcelTools
{
    public partial class ThisAddIn
    {
        #region Private Members

        private ServicesProvider _serviceProvider;
        //private SplashScreen splashScreen;

        #endregion

        #region Private Helpers

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            CultureService.SetCulture();

            COMAddIn addIn = Globals.ThisAddIn.Application.COMAddIns.Item(nameof(XcelTools));

            XcelToolsApplication application = new XcelToolsApplication(addIn);
            application.Start();
        }

        private void OnTimerElapsed(object sender, ElapsedEventArgs e)
        {
            
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {

        }

        protected override object RequestComAddInAutomationService()
        {
            return _serviceProvider ?? (_serviceProvider = new ServicesProvider());
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += new EventHandler(ThisAddIn_Startup);
            Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
