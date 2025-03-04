using Microsoft.Office.Interop.Excel;

using System.Runtime.InteropServices;

namespace XcelTools.Xtractor.Interfaces
{
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISecurityService
    {
        void LockWorksheet(Worksheet worksheet);
        void UnlockWorksheet(Worksheet worksheet);
    }
}