using System;
using System.Runtime.InteropServices;
using System.Windows;

using XcelTools.Xtractor.Interfaces;

using Excel = Microsoft.Office.Interop.Excel;

namespace XcelTools.Xtractor.Services
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class SecurityService : ISecurityService
    {
        #region Default Constructor

        public SecurityService()
        {
            Password = "EngineeredProgramming";
        }

        #endregion

        #region Private Properties

        private string Password { get; set; }

        #endregion

        public void LockWorksheet(Excel.Worksheet worksheet)
        {
            try
            {
                worksheet.Protect(Password, UserInterfaceOnly: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Following error encountered! {Environment.NewLine}{ex.Message}", "Lock Worksheet");
            }
        }

        public void UnlockWorksheet(Excel.Worksheet worksheet)
        {
            worksheet.Unprotect(Password);
        }
    }
}
