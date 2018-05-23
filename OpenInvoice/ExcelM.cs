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
        }
    }
}
