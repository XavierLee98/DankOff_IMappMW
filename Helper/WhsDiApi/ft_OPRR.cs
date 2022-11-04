using IMAppSapMidware_NetCore.Models.SAPModels;
using System;
using System.Collections.Generic;
using System.Data;
using System.Runtime.InteropServices;
using System.Text;

namespace IMAppSapMidware_NetCore.Helper.SQL
{
    class ft_OPRR
    {
        public static string LastSAPMsg { get; set; } = string.Empty;
        // added by jonny to track error when unexpected error
        // 20210411
        static string currentKey = string.Empty;
        static string currentStatus = string.Empty;
        static string CurrentDocNum = string.Empty;

        static void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }
        public static void Post()
        {
            DataTable dt = null;

            string sapdb = "";
            try
            {
                dt = ft_General.LoadData("LoadOPRR_sp");
                //dtDetails = ft_General.LoadData("LoadOWTQDetails_sp");
                string failed_status = "ONHOLD";
                string success_status = "SUCCESS";
                string tablename = "OPRR";
                string docnum = "";
                string docEntry = "";
                int cnt = 0;
                int retcode = 0;

                if (dt.Rows.Count > 0)
                {
                    SAPParam par = SAP.GetSAPUser();
                    SAPCompany sap = SAP.getSAPCompany(par);

                    if (!sap.connectSAP()) throw new Exception(sap.errMsg);

                    string key = dt.Rows[0]["key"].ToString();
                    // added by jonny to track error when unexpected error
                    // 20210411
                    currentKey = key;
                    currentStatus = failed_status;


                    SAPbobsCOM.Recordset rc = null;
                    SAPbobsCOM.Documents oDoc = null;// (SAPbobsCOM.StockTransfer)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (!sap.oCom.InTransaction)
                            sap.oCom.StartTransaction();

                        if (cnt > 0)
                        {
                            oDoc.Lines.Add();
                            oDoc.Lines.SetCurrentLine(cnt);

                            if (key == dt.Rows[i]["key"].ToString()) goto details;

                            retcode = oDoc.Add();// Add record 
                            if (retcode != 0) // if error
                            {
                                if (sap.oCom.InTransaction)
                                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
                                Log($"{key }\n {failed_status }\n { message } \n");
                                ft_General.UpdateStatus(key, failed_status, message, "");
                            }
                            else
                            {
                                sap.oCom.GetNewObjectCode(out docEntry);
                                docnum = ft_General.GetDocNum(sap.oCom, tablename, docEntry);
                                if (sap.oCom.InTransaction)
                                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                Log($" {key }\n {success_status }\n  { docnum } \n");
                                ft_General.UpdateStatus(key, success_status, "", docnum);
                            }

                            cnt = 0;
                            if (oDoc != null) Marshal.ReleaseComObject(oDoc);
                            if (rc != null) Marshal.ReleaseComObject(rc);
                            oDoc = null;
                            rc = null;
                        }

                        if (!sap.oCom.InTransaction)
                            sap.oCom.StartTransaction();

                        oDoc = (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oGoodsReturnRequest);
                        rc = (SAPbobsCOM.Recordset)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        oDoc.CardCode = dt.Rows[i]["cardcode"].ToString();
                        oDoc.DocDate = DateTime.Parse(dt.Rows[i]["docdate"].ToString());
                        oDoc.TaxDate = DateTime.Parse(dt.Rows[i]["docdate"].ToString());
                        oDoc.DocDueDate = DateTime.Parse(dt.Rows[i]["docdate"].ToString());
                        if (dt.Rows[i]["ref2"].ToString() != "")
                            oDoc.Reference2 = dt.Rows[i]["ref2"].ToString();
                        if (dt.Rows[i]["comments"].ToString() != "")
                            oDoc.Comments = dt.Rows[i]["comments"].ToString();
                        if (dt.Rows[i]["jrnlmemo"].ToString() != "")
                            oDoc.JournalMemo = dt.Rows[i]["jrnlmemo"].ToString();
                        if (dt.Rows[i]["numatcard"].ToString() != "")
                            oDoc.NumAtCard = dt.Rows[i]["numatcard"].ToString();


                        details:
                        oDoc.Lines.ItemCode = dt.Rows[i]["itemcode"].ToString();
                        oDoc.Lines.Quantity = double.Parse(dt.Rows[i]["quantity"].ToString());
                        oDoc.Lines.WarehouseCode = dt.Rows[i]["whscode"].ToString();

                        // add in the pur in remark 
                        // 20210403
                        var itemDetails = dt.Rows[i]["remarks"].ToString();
                        if (!string.IsNullOrWhiteSpace(itemDetails))
                        {
                            oDoc.Lines.ItemDetails = itemDetails;
                        }

                        if (int.Parse(dt.Rows[i]["baseentry"].ToString()) > 0)
                        {
                            oDoc.Lines.BaseEntry = int.Parse(dt.Rows[i]["baseentry"].ToString());
                            oDoc.Lines.BaseLine = int.Parse(dt.Rows[i]["baseline"].ToString());
                            oDoc.Lines.BaseType = int.Parse(dt.Rows[i]["basetype"].ToString());
                            if (dt.Rows[i]["whscode"].ToString() != "")
                                oDoc.Lines.WarehouseCode = dt.Rows[i]["whscode"].ToString();
                            else
                            {
                                string sourceTableName = dt.Rows[i]["sourcetablename"].ToString();
                                string sourceLineTableName = sourceTableName.Substring(1, 3) + "1";
                                rc.DoQuery("select * from " + sourceLineTableName + " where docentry = " + int.Parse(dt.Rows[i]["baseentry"].ToString()) + " and linenum = " + int.Parse(dt.Rows[i]["baseline"].ToString()));
                                if (rc.RecordCount > 0)
                                {
                                    oDoc.Lines.WarehouseCode = rc.Fields.Item("whscode").Value.ToString();
                                }
                            }
                        }
                        else
                        {
                            if (dt.Rows[i]["whscode"].ToString() != "")
                                oDoc.Lines.WarehouseCode = dt.Rows[i]["whscode"].ToString();
                        }

                        key = dt.Rows[i]["key"].ToString();
                        cnt++;
                    }
                    retcode = oDoc.Add();
                    if (retcode != 0)
                    {
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
                        Log($"{key }\n {failed_status }\n { message } \n");
                        ft_General.UpdateStatus(key, failed_status, message, "");
                    }
                    else
                    {
                        sap.oCom.GetNewObjectCode(out docEntry);
                        docnum = ft_General.GetDocNum(sap.oCom, tablename, docEntry);

                        // added by jonny to track error when unexpected error
                        // 20210411
                        CurrentDocNum = docnum;

                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        Log($" {key }\n {success_status }\n  { docnum } \n");
                        ft_General.UpdateStatus(key, success_status, "", docnum);
                    }

                    if (oDoc != null) Marshal.ReleaseComObject(oDoc);
                    if (rc != null) Marshal.ReleaseComObject(rc);
                    oDoc = null;
                    rc = null;
                }
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
                ft_General.UpdateError("OPRR", ex.Message);

                // added by jonny to track error when unexpected error
                // 20210411
                Log($"{currentKey }\n {currentStatus }\n { ex.Message } \n");
                ft_General.UpdateStatus(currentKey, currentStatus, ex.Message, CurrentDocNum);
            }
            finally
            {
                dt = null;
            }
        }
    }
}
