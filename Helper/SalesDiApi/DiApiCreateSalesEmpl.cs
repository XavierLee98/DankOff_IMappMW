using Dapper;
using Microsoft.Data.SqlClient;
using IMAppSapMidware_NetCore.Models.SalesEmployee;
using SAPbobsCOM;
using System;

namespace IMAppSapMidware_NetCore.Helper.DiApi
{
    public class DiApiCreateSalesEmpl
    {   
        public ErpPropertyHelper ErpProperty { get; set; } = null;
        public string LastSAPMsg { get; set; } = string.Empty;
        public string Erp_DBConnStr { get; set; } = string.Empty;        
        public string SalesEmpCode { get; set; } = string.Empty;
        public SalesEmployee SalesPerson { get; set; } = null;
        
        Company sapCompany { get; set; }

        void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }

        public bool CreateSalesEmployee()
        {
            string modName = $"[CreateSalesEmployee][{SalesPerson.DisplayName}]";
            try
            {
                if (SalesPerson == null) return false;
                // connect the diapi company from class
                // maintain in one place
                // replace if (ConnectDI() != 0) return false;  // connect the server if fail then return                                
                var di = new DiApiUtilities { ErpProperty = this.ErpProperty };
                if (di.ConnectDI() != 0)
                {
                    Log($"{modName}\n{di.LastErrorMessage}");
                    return false;
                }
                sapCompany = di.SapCompany;

                var sales = (SalesPersons)sapCompany.GetBusinessObject(BoObjectTypes.oSalesPersons);
                sales.Active = BoYesNoEnum.tYES;

                // only one row reading
                var salesEmpName = SalesPerson.UserIdName;
                sales.SalesEmployeeName = FormatValue(salesEmpName, 155); // 155 chars                      

                var remarks = "Created by App SuperAdmin";
                sales.Remarks = FormatValue(remarks, 50);

                if (sales.Add() == 0) // success added
                {
                    SalesEmpCode = GetSalesEmpCode(salesEmpName); //sales.SalesEmployeeCode.ToString();
                    return true;
                }

                // else                                    
                Log($"{sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
                return false;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return false;
            }
        }

        string FormatValue(string val, int requireLength)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(val)) return string.Empty;                
                if (val.Length > requireLength) return val.Substring(0, requireLength);               
                return val;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return string.Empty;
            }
        }

        /// <summary>
        /// query database to get the sales employee name
        /// </summary>
        /// <param name="salesEmpName"></param>
        /// <returns></returns>
        string GetSalesEmpCode(string salesEmpName)
        {
            try
            {
                var result =  new SqlConnection(this.Erp_DBConnStr)
                    .ExecuteScalar<string>("SELECT SlpCode FROM OSLP WHERE SlpName = @salesEmpName", new { salesEmpName });

                return result;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return string.Empty;
            }
        }
    }
}
