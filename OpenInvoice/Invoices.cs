using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Metro
{
    public class Invoices
    {
        public static void Open(string invoiceNum)
        {
            openInvoice(invoiceNum);
        }

        public static void Open(string[] invoiceNums)
        {
            for(int i = 0; i < invoiceNums.Length; i++)
            {
                openInvoice(invoiceNums[i]);
            }
        }

        private static void openInvoice(string invoiceNum)
        {
            DirectoryInfo _dInfo = new DirectoryInfo(@"\\METRO-GP1\Dynamics\MESSNGER\Archive");
            string[] _dirs = Directory.GetDirectories(@"\\METRO-GP1\Dynamics\MESSNGER\", "20*Invoices");

            FileInfo[] filesInDir = _dInfo.GetFiles("in-*" + "-" + invoiceNum + "-" + invoiceNum + "-" + "*");

            int iCounter = _dirs.Length - 1;

            while(filesInDir.Length == 0)
            {
                if (iCounter < 0) break;

                DirectoryInfo dInfo = new DirectoryInfo(_dirs[iCounter]);
                filesInDir = dInfo.GetFiles("in-*" + "-" + invoiceNum + "-" + invoiceNum + "-" + "*");

                iCounter--;
            }

            try
            {
                System.Diagnostics.Process.Start(filesInDir[0].FullName);
            }
            catch (Exception)
            {
                MessageBox.Show("Error Finding and Opening Invoice. Please use Liason.\nInvoice Number: " + invoiceNum);
            }
        }
    }
}
