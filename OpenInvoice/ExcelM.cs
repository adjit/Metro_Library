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

        public static void Export(DataTable dt)
        {
            Excel.Application exApp = new Excel.Application();
            Excel.Workbooks wbs = exApp.Workbooks;
            Excel._Workbook wb = null;

            wb = (Excel._Workbook)(wbs.Add(Excel.XlWBATemplate.xlWBATWorksheet));

            Excel.Worksheet sheet = wb.ActiveSheet;

            for(int r = 0; r <= dt.Rows.Count; r++)
            {
                for(int c = 0; c < dt.Columns.Count; c++)
                {
                    if (r == 0) sheet.Cells[r + 1, c + 1].Value2 = "'" + dt.Columns[c].ColumnName;
                    else
                    {
                        sheet.Cells[r + 1, c + 1].Value2 = "'" + dt.Rows[r - 1][c];
                    }
                }
            }

            exApp.Visible = true;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.FinalReleaseComObject(sheet);
            Marshal.FinalReleaseComObject(wb);
            Marshal.FinalReleaseComObject(wbs);
            Marshal.FinalReleaseComObject(exApp);
        }
    }
}
