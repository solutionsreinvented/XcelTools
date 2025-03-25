
using Excel = Microsoft.Office.Interop.Excel;

namespace XcelTools.Services
{
    public class ExcelCommonServices
    {
        public static void PopulateListInACell(Excel.Range range, string listFormula)
        {
            range.Select();

            Excel.Validation rngValidation = range.Validation;

            rngValidation.Delete();
            rngValidation.Add(Excel.XlDVType.xlValidateList, AlertStyle: Excel.XlDVAlertStyle.xlValidAlertStop, Operator: Excel.XlFormatConditionOperator.xlBetween, Formula1: listFormula);
            rngValidation.IgnoreBlank = true;
            rngValidation.InCellDropdown = true;
            rngValidation.InputTitle = string.Empty;
            rngValidation.ErrorTitle = string.Empty;
            rngValidation.InputMessage = string.Empty;
            rngValidation.ErrorMessage = string.Empty;
            rngValidation.ShowInput = true;
            rngValidation.ShowError = true;
        }
    }
}
