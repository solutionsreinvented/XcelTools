using System;
using System.Timers;

using XcelTools.Licensing.Views;
using XcelTools.Models;

namespace XcelTools
{
    public partial class ThisAddIn
    {
        private ServicesProvider _serviceProvider;
        private SplashScreen splashScreen;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            
            Timer timer = new Timer(TimeSpan.FromSeconds(1).TotalMilliseconds);
            timer.Elapsed += OnTimerElapsed;
            timer.Start();
            splashScreen = new SplashScreen();
            splashScreen.Show();
            timer.Stop();
        }

        private void OnTimerElapsed(object sender, ElapsedEventArgs e)
        {
            if (splashScreen != null)
            {
                splashScreen.Close();
            }
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
