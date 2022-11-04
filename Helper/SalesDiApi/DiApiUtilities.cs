using SAPbobsCOM;
using System;

namespace IMAppSapMidware_NetCore.Helper.DiApi
{
    public class DiApiUtilities
    {
        public string LastErrorMessage { get; set; }
        public Company SapCompany { get; set; }
        public ErpPropertyHelper ErpProperty { get; set; }

        public int ConnectDI()
        {   // for connecting the app to the SAP company object = Di API

            if (Program.SapCompany != null && Program.SapCompany.Connected)
            {
                SapCompany = Program.SapCompany;
                return 0;
            }

            int result = -1;
            SapCompany = new Company();
            SapCompany.Server = ErpProperty.Server;        // MALP03\MSSQLSERVERV16
            SapCompany.DbServerType = ErpProperty.BoDataServerTypes;   // BoDataServerTypes.dst_MSSQL2008;
            SapCompany.CompanyDB = ErpProperty.CompanyDB;//  sapDatabase;       //// sapDatabase; //"DB_TestConnection";

            switch (ErpProperty.PortNumber) // defaultPortNum)
            {
                case "30000":
                    {
                        SapCompany.LicenseServer = $"{ErpProperty.LicenseServer}:{ErpProperty.PortNumber}";  // $"{svrMachineName}:{defaultPortNum}";
                        break;
                    }
                default: // any port number 
                    {
                        SapCompany.SLDServer = $"{ErpProperty.LicenseServer}:{ErpProperty.PortNumber}"; //$"{svrMachineName}:{defaultPortNum}";
                        break;
                    }
            }

            SapCompany.UserName = ErpProperty.UserName; // sapUserName; /// sapUserName; //"manager";
            SapCompany.Password = ErpProperty.Password; // sapPassword; ///sapPassword;//"P@ssw0rd";
            try
            {
                result = SapCompany.Connect();
                if (result == 0 || result == -116)
                {
                    Program.SapCompany = SapCompany; // assign back to the global company object                                                                    
                    result = 0;
                    return result;
                }

                Log($"BpLead Di Connect {result}\n{SapCompany.GetLastErrorCode()}\n{SapCompany.GetLastErrorDescription()}");
                return result;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return result;
            }
        }

        void Log(string message)
        {
            LastErrorMessage += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }
    }
}
