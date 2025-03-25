using Microsoft.Office.Interop.Excel;

namespace XcelTools.Extensions
{
    public static class WorksheetExtensions
    {
        public static string RangeAddress(this Worksheet sheet, int sRow, int sCol, int eRow, int eCol)
        {
            return "=" + sheet.Name + "!" + sheet.Range[sheet.Cells[sRow, sCol], sheet.Cells[eRow, eCol]].Address;
        }
    }
}
