using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace IMAppSapMidware_NetCore.Helper.SQL
{
    class ft_General
    {
        public static void UpdateError(string module, string errMsg)
        {
            DataTable dt = new DataTable();
            string spName = "UpdateError_sp";

            dt = DAC.ExecuteDataTable(spName,
                DAC.Parameter("@MODULE", module),
                DAC.Parameter("@ERRMSG", errMsg));
        }
        public static DataTable LoadData(string spname)
        {
            DataTable dt = new DataTable();

            dt = DAC.ExecuteDataTable(spname);

            return dt;
        }
        public static DataTable LoadDataByRequest(string spname, string request)
        {
            DataTable dt = new DataTable();

            dt = DAC.ExecuteDataTable(spname,
                  DAC.Parameter("@REQUEST", request));

            return dt;
        }
        public static DataTable LoadDataByGuid(string spname, string guid)
        {
            DataTable dt = new DataTable();

            dt = DAC.ExecuteDataTable(spname,
                  DAC.Parameter("@GUID", guid));

            return dt;
        }

        public static void UpdateStatus(string key, string status, string errMsg, string docnum)
        {
            DataTable dt = new DataTable();
            string spName = "UpdateStatus_sp";

            dt = DAC.ExecuteDataTable(spName,
                DAC.Parameter("@KEY", key),
                DAC.Parameter("@STATUS", status),
                DAC.Parameter("@ERRMSG", errMsg),
                DAC.Parameter("@DOCNUM", docnum));
        }
        public static string GetDocNum(SAPbobsCOM.Company oCom, string tablename, string docentry)
        {
            string docnum = "";
            SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rc.DoQuery("select * from " + tablename + " where docentry =" + docentry);

            if (rc.RecordCount > 0) docnum = rc.Fields.Item("docnum").Value.ToString();

            return docnum;
        }
        public static DataTable LoadBinBatchSerial(string key, string itemcode)
        {
            DataTable dt = new DataTable();
            string spName = "GetBinBatchSerial_sp";

            dt = DAC.ExecuteDataTable(spName,
                DAC.Parameter("@KEY", key),
                DAC.Parameter("@ITEMCODE", itemcode));

            return dt;
        }
    }
}
