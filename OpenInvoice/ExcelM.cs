using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Metro
{
    /*      The ExcelM class of the Metro package will be responsible for
     *      exporting data to an Excel document (provided Excel is installed) 
     */
    class ExcelM
    {

        public static void Export(DataTable dt)
        {
            Excel.Application exApp = new Excel.Application();
            Excel._Workbook wb = null;

            wb = (Excel._Workbook)(exApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet));

            Excel.Worksheet sheet = wb.ActiveSheet;

            for(int r = 0; r < dt.Rows.Count; r++)
            {
                for(int c = 0; c < dt.Columns.Count; c++)
                {
                    sheet.Cells[r + 1, c + 1].Value2 = dt.Rows[r][c];
                }
            }

            exApp.Visible = true;
        }
    }
}
