using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Metro
{
    /*      The ExcelM class of the Metro package will be responsible for
     *      exporting data to an Excel document (provided Excel is installed) 
     */
    public class ExcelM
    {

        private static Excel.Application exApp;
        private static Excel.Workbooks wbs;
        private static Excel._Workbook wb;
        private static Excel.Worksheet sheet;

        private static void _initializeExcel()
        {
            exApp = new Excel.Application();
            wbs = exApp.Workbooks;
            wb = (Excel._Workbook)(wbs.Add(Excel.XlWBATemplate.xlWBATWorksheet));
            sheet = wb.ActiveSheet;
        }

        private static void _cleanObjects()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(sheet);
            Marshal.FinalReleaseComObject(wb);
            Marshal.FinalReleaseComObject(wbs);
            Marshal.FinalReleaseComObject(exApp);
        }

        public static void Export(DataTable dt)
        {
            _Export(dt);

            exApp.Visible = true;

            _cleanObjects();
        }

        public static void Export(DataTable dt, Boolean showExcel)
        {
            _Export(dt);

            if (showExcel)
                exApp.Visible = true;

            _cleanObjects();
        }

        public static void Export(DataTable dt, Boolean showExcel, string saveAs)
        {
            _Export(dt);

            /*wb.SaveAs(saveAs, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing);
                            */
            if (showExcel)
                exApp.Visible = true;
            else
            {
                exApp.DisplayAlerts = false;
                wb.Close(SaveChanges:true, Filename:saveAs);
                wbs.Close();
                exApp.Quit();
            }
            _cleanObjects();
        }

        private static void _Export(DataTable dt)
        {
            _initializeExcel();

            for (int r = 0; r <= dt.Rows.Count; r++)
            {
                for (int c = 0; c < dt.Columns.Count; c++)
                {
                    if (r == 0) sheet.Cells[r + 1, c + 1].Value2 = "'" + dt.Columns[c].ColumnName;
                    else
                    {
                        string colName = dt.Columns[c].ColumnName.ToUpper();

                        if (colName.Contains("PRICE") || colName.Contains("COST"))
                        {
                            sheet.Cells[r + 1, c + 1].Value2 = dt.Rows[r - 1][c];
                            sheet.Cells[r + 1, c + 1].NumberFormat = "$#,##0.00";
                        }
                        else if (colName.Contains("QTY"))
                        {
                            sheet.Cells[r + 1, c + 1].Value2 = dt.Rows[r - 1][c];
                        }
                        else
                            sheet.Cells[r + 1, c + 1].Value2 = "'" + dt.Rows[r - 1][c];
                    }
                }
            }
        }
    }
}
