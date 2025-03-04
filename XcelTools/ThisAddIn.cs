using System;
using XcelTools.Models;

using XcelTools.XpLore.Views;
using XcelTools.Xtractor.Models.Sections;
using XcelTools.Xtractor.Services;

namespace XcelTools
{
    public partial class ThisAddIn
    {
        private ServicesProvider _serviceProvider;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            //XploreWindow xploreWindow = new XploreWindow();

            //xploreWindow.Show();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {

        }

        protected override object RequestComAddInAutomationService()
        {
            return _serviceProvider ?? (_serviceProvider = new ServicesProvider());
        }

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
