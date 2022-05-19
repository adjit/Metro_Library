using System;
using System.Data;
using System.Net;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Avalara.AvaTax.RestClient;

namespace Metro
{
    public class Avalara
    {
        private static AvaTaxClient avClient;
        private static int METSALES = 196080;
        private static bool connectionStatus;
        private static PingResultModel pingResult;

        public static DataTable ExemptionLookup(string customer)
        {
            DataTable responseTable = new DataTable();

            string connection, username, login;
            try
            {
                connection = System.IO.File.ReadAllText(@"\\METRO-FILE1\Metropolitan Sales Docs\1-Deployment\avConnect\avConnection");
                var connections = connection.Split(';');
                username = connections[0];
                login = connections[1];
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("Unable to get connection string for Avalara.");
                responseTable.Columns.Add("ERROR");
                responseTable.Rows.Add(new object[] { "ERROR : Cannot connect to Avalara" });
                return responseTable;
            }

            try
            {
                avClient = new AvaTaxClient("MetroTools", "1.0", Environment.MachineName, AvaTaxEnvironment.Production)
                                .WithSecurity(username, login);

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

                pingResult = avClient.Ping();
            }
            catch (Exception ex)
            {
                //string errorMessage = $"{ex.Message} ---> {ex.InnerException.Message} ---> ";
                StringBuilder errorMessage = new StringBuilder(ex.Message);
                errorMessage.AppendLine();
                Exception currentEx = ex;
                for(int i = 0; i < 2; i++)
                {
                    if(currentEx.InnerException != null)
                    {
                        currentEx = currentEx.InnerException;
                        errorMessage.AppendLine($"---> {currentEx.Message}");
                    }
                }

                System.Windows.Forms.MessageBox.Show($"Unable to connect to Avalara.\n{errorMessage}");
                responseTable.Columns.Add("ERROR");
                responseTable.Rows.Add(new object[] { ex.Message });
                return responseTable;
            }

            connectionStatus = (bool)pingResult.authenticated;

            if (connectionStatus)
            {
                //var customerResult = avClient.ListCertificatesForCustomer(METSALES, customer, null, null, null, null, "exposureZone ASC");
                var customerResult = avClient.ListCertificatesForCustomer(METSALES, customer, null, null, null, null, null);

                responseTable.Columns.Add("Exposure Zone", typeof(string));
                responseTable.Columns.Add("Exempt Reason", typeof(string));
                responseTable.Columns.Add("Expires", typeof(DateTime));

                if (customerResult.count == 0)
                {
                    responseTable.Rows.Add("NO RESALE", "Not Exempt", DateTime.Now);
                }
                else
                {
                    for (int i = 0; i < customerResult.count; i++)
                    {
                        var thisVal = customerResult.value[i];
                        string exZone = thisVal.exposureZone.name;
                        string exReason = thisVal.exemptionReason.name;
                        bool validity = (bool)thisVal.valid;
                        DateTime signDate = thisVal.signedDate;
                        DateTime expiryDate = thisVal.expirationDate;

                        responseTable.Rows.Add(exZone, exReason, expiryDate);
                    }
                }

            }
            else return null;

            return responseTable;
        }
    }
}
