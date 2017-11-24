using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace HS.ExcelExt
{
    public static class SheetExt
    {
        public static void CleanPassword(this Excel.Worksheet sheet)
        {
            #region 一个小功能
            sheet.Protect
            (
                DrawingObjects: Office.MsoTriState.msoTrue, 
                Contents: Office.MsoTriState.msoTrue, 
                Scenarios: Office.MsoTriState.msoTrue, 
                AllowFiltering: Office.MsoTriState.msoTrue, 
                AllowUsingPivotTables: Office.MsoTriState.msoTrue
            );

            sheet.Protect
            (
                DrawingObjects: Office.MsoTriState.msoTrue, 
                Contents: Office.MsoTriState.msoTrue, 
                Scenarios: Office.MsoTriState.msoTrue, 
                AllowFiltering: Office.MsoTriState.msoTrue, 
                AllowUsingPivotTables: Office.MsoTriState.msoTrue
            );

            sheet.Protect
            (
                DrawingObjects: Office.MsoTriState.msoTrue, 
                Contents: Office.MsoTriState.msoTrue, 
                Scenarios: Office.MsoTriState.msoTrue, 
                AllowFiltering: Office.MsoTriState.msoTrue, 
                AllowUsingPivotTables: Office.MsoTriState.msoTrue
            );

            sheet.Protect
            (
                DrawingObjects: Office.MsoTriState.msoTrue, 
                Contents: Office.MsoTriState.msoTrue, 
                Scenarios: Office.MsoTriState.msoTrue, 
                AllowFiltering: Office.MsoTriState.msoTrue, 
                AllowUsingPivotTables: Office.MsoTriState.msoTrue
            );

            sheet.Unprotect();
            #endregion
        }
    }
}
