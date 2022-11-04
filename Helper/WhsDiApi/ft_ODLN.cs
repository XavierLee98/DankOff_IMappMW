using Dapper;
using IMAppSapMidware_NetCore.Models.SAPModels;
using Microsoft.Data.SqlClient;
using System;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace IMAppSapMidware_NetCore.Helper.SQL
{
    class ft_ODLN
    {
        public static string LastSAPMsg { get; set; } = string.Empty;

        // added by jonny to track error when unexpected error
        // 20210411
        static string currentKey = string.Empty;
        static string currentStatus = string.Empty;
        static string CurrentDocNum = string.Empty;
        public static string Erp_DBConnStr { get; set; } = string.Empty;
        static bool isOtherUOM = false;
        static UOMConvert unit = null;

        static void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }
        public async static void Post()
        {
            DataTable dt = null;
            DataTable dtDetails = null;
            DataTable dtBin = null;
            string request = "Create DO";

            string sapdb = "";
            try
            {
                dt = ft_General.LoadData("LoadODLN_sp");
                dtDetails = ft_General.LoadDataByRequest("LoadDetails_sp", request);
                dtBin = ft_General.LoadDataByRequest("LoadBinDetails_sp", request);
                string failed_status = "ONHOLD";
                string success_status = "SUCCESS";
                string tablename = "ODLN";
                string docnum = "";
                string docEntry = "";
                int cnt = 0, bin_cnt = 0, batch_cnt = 0, serial_cnt = 0, batchbin_cnt = 0;
                int retcode = 0;

                if (dt.Rows.Count > 0)
                {
                    SAPParam par = SAP.GetSAPUser();
                    SAPCompany sap = SAP.getSAPCompany(par);

                    if (!sap.connectSAP())
                    {
                        Log($"{sap.errMsg}");
                        throw new Exception(sap.errMsg);

                    }

                    string key = dt.Rows[0]["key"].ToString();
                    // added by jonny to track error when unexpected error
                    // 20210411
                    currentKey = key;
                    currentStatus = failed_status;

                    SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Documents oDoc = null;// (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
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

                                // added by jonny to track error when unexpected error
                                // 20210411
                                CurrentDocNum = docnum;


                                if (sap.oCom.InTransaction)
                                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                Log($"{key }\n {success_status }\n  { docnum } \n");
                                ft_General.UpdateStatus(key, success_status, "", docnum);
                            }
                            cnt = 0;
                            if (oDoc != null) Marshal.ReleaseComObject(oDoc);
                            oDoc = null;
                        }

                        if (!sap.oCom.InTransaction)
                            sap.oCom.StartTransaction();

                        oDoc = (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);

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
                        // added by KX
                        // 20210412
                        oDoc.Lines.ItemCode = dt.Rows[i]["itemcode"].ToString();
                        oDoc.Lines.UoMEntry = int.Parse(dt.Rows[i]["UomEntry"].ToString());
                        oDoc.Lines.Quantity = double.Parse(dt.Rows[i]["quantity"].ToString());

                        if(dt.Rows[i]["UomEntry"].ToString() != "Manual")
                        {
                            isOtherUOM = true;
                            unit = await GetUOMUnit(dt.Rows[i]["UomCode"].ToString());
                            if (unit == null) throw new Exception("UOM Unit is null.");
                        }

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
                                rc.DoQuery("select * from rdr1 where docentry = " + int.Parse(dt.Rows[i]["baseentry"].ToString()) + " and linenum = " + int.Parse(dt.Rows[i]["baseline"].ToString()));
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

                        //DataTable dtBinBatchSerial = ft_General.LoadBinBatchSerial(dt.Rows[i]["key"].ToString(), dt.Rows[i]["itemcode"].ToString());
                        DataRow[] dr = dtDetails.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "' " +
                            " and baseentry=" + dt.Rows[i]["baseentry"].ToString() + " and basetype ='" + dt.Rows[i]["basetype"].ToString() + "'");
                        if (dr.Length > 0)
                        {
                            for (int x = 0; x < dr.Length; x++)
                            {
                                if (dr[x]["batchnumber"].ToString() != "")
                                {
                                    // added by KX
                                    // 20210412
                                    //var numberinsale = GetNumInSale(itemcode);
                                    //if (numberinsale == -1)
                                    //{
                                    //    numberinsale = double.Parse(decimal.Parse(dr[x]["quantity"].ToString()).ToString());
                                    //}
                                    if (batch_cnt > 0) oDoc.Lines.BatchNumbers.Add();
                                    oDoc.Lines.BatchNumbers.SetCurrentLine(batch_cnt);
                                    oDoc.Lines.BatchNumbers.BatchNumber = dr[x]["batchnumber"].ToString();

                                    // added by KX
                                    // 20210412
                                    //oDoc.Lines.BatchNumbers.Quantity = oDoc.Lines.Quantity * numberinsale;
                                    // oDoc.Lines.BatchNumbers.Quantity = double.Parse(decimal.Parse(dr[x]["quantity"].ToString()).ToString());
                                    if (!isOtherUOM)
                                        oDoc.Lines.BatchNumbers.Quantity = double.Parse(dr[x]["quantity"].ToString());
                                    else
                                    {
                                        oDoc.Lines.BatchNumbers.Quantity = ConvertUOMQuantity(double.Parse(dr[x]["quantity"].ToString()));
                                    }
                                    oDoc.Lines.BatchNumbers.ManufacturerSerialNumber = dr[x]["batchattr1"].ToString();

                                    if (dr[x]["batchadmissiondate"].ToString() != "")
                                        oDoc.Lines.BatchNumbers.AddmisionDate = DateTime.Parse(dr[x]["batchadmissiondate"].ToString());

                                    DataRow[] drBin = dtBin.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() +
                                        "' and Batchnumber ='" + dr[x]["batchnumber"].ToString() + "' " +
                                         " and baseentry=" + dt.Rows[i]["baseentry"].ToString() + " and basetype ='" + dt.Rows[i]["basetype"].ToString() + "'");

                                    if (drBin.Length > 0)
                                    {
                                        for (int y = 0; y < drBin.Length; y++)
                                        {
                                            if (batchbin_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                            oDoc.Lines.BinAllocations.SetCurrentLine(batchbin_cnt);
                                            oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBin[y]["binabsentry"].ToString());
                                            if (!isOtherUOM)
                                                oDoc.Lines.BinAllocations.Quantity = double.Parse(drBin[y]["quantity"].ToString());
                                            else
                                            {
                                                oDoc.Lines.BinAllocations.Quantity = ConvertUOMQuantity(double.Parse(drBin[y]["quantity"].ToString()));
                                            }
                                            //oDoc.Lines.BinAllocations.Quantity = double.Parse(drBin[y]["quantity"].ToString());
                                            oDoc.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = batch_cnt;
                                            batchbin_cnt++;
                                        }
                                    }

                                    batch_cnt++;
                                }
                                else if (dr[x]["serialnumber"].ToString() != "")
                                {
                                    if (serial_cnt > 0) oDoc.Lines.SerialNumbers.Add();
                                    oDoc.Lines.SerialNumbers.SetCurrentLine(serial_cnt);
                                    oDoc.Lines.SerialNumbers.InternalSerialNumber = dr[x]["serialnumber"].ToString();
                                    DataRow[] drBin = dtBin.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() +
                                        "' and serialnumber ='" + dr[x]["serialnumber"].ToString() + "' " +
                                         " and baseentry=" + dt.Rows[i]["baseentry"].ToString() + " and basetype ='" + dt.Rows[i]["basetype"].ToString() + "'");

                                    if (drBin.Length > 0)
                                    {
                                        if (serial_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                        oDoc.Lines.BinAllocations.SetCurrentLine(serial_cnt);
                                        oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBin[0]["binabsentry"].ToString());
                                        oDoc.Lines.BinAllocations.Quantity = double.Parse(drBin[0]["quantity"].ToString());
                                        oDoc.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = serial_cnt;
                                    }

                                    serial_cnt++;
                                }
                                else
                                {
                                    DataRow[] drBin = dtBin.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "' " +
                                         " and baseentry=" + dt.Rows[i]["baseentry"].ToString() + " and basetype ='" + dt.Rows[i]["basetype"].ToString() + "'");

                                    if (drBin.Length > 0)
                                    {
                                        for (int y = 0; y < drBin.Length; y++)
                                        {
                                            if (bin_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                            oDoc.Lines.BinAllocations.SetCurrentLine(bin_cnt);
                                            oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBin[y]["binabsentry"].ToString());
                                            //oDoc.Lines.BinAllocations.Quantity = double.Parse(drBin[y]["quantity"].ToString());
                                            //oDoc.Lines.BinAllocations.BaseLineNumber = cnt;
                                            if (!isOtherUOM)
                                                oDoc.Lines.BinAllocations.Quantity = double.Parse(drBin[y]["quantity"].ToString());
                                            else
                                            {
                                                oDoc.Lines.BinAllocations.Quantity = ConvertUOMQuantity(double.Parse(drBin[y]["quantity"].ToString()));
                                            }

                                            bin_cnt++;
                                        }
                                    }
                                }
                            }
                            bin_cnt = 0;
                            serial_cnt = 0;
                            batch_cnt = 0;
                            batchbin_cnt = 0;
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
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        Log($"{key }\n {success_status }\n  { docnum } \n");
                        ft_General.UpdateStatus(key, success_status, "", docnum);
                    }

                    if (oDoc != null) Marshal.ReleaseComObject(oDoc);
                    oDoc = null;
                }
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
                ft_General.UpdateError("ODLN", ex.Message);

                Log($"{currentKey }\n {currentStatus }\n { ex.Message } \n");
                ft_General.UpdateStatus(currentKey, currentStatus, ex.Message, CurrentDocNum);
            }
            finally
            {
                dt = null;
                dtDetails = null;
                dtBin = null;
            }
        }

        // added by KX
        // 20210412
        static double GetNumInSale(string ItemCode)
        {
            Erp_DBConnStr= Program._DbErpConnStr;
            var parameter = new { ItemCode };

            string query = "Select NumInSale from OITM where ItemCode = @ItemCode";
            // sql to query the OITM table to get the num in buy
            using (var conn = new SqlConnection(Erp_DBConnStr))
            {
                return conn.QuerySingle<double>(query, parameter);
            }
        }

        static async Task<UOMConvert> GetUOMUnit(string FromUomCode)
        {
            try
            {
                var conn = new System.Data.SqlClient.SqlConnection(Program._DbErpConnStr);
                string query = $"SELECT T1.AltQty [FromUnit], T1.BaseQty [ToUnit] FROM OUOM T0 " +
                               $"INNER JOIN UGP1 T1 on T1.UomEntry = T0.UomEntry " +
                               $"WHERE T0.UomCode = @UomCode";

                return conn.Query<UOMConvert>(query, new { UomCode = FromUomCode }).FirstOrDefault();
            }
            catch (Exception e)
            {
                LastSAPMsg = e.ToString();
                return null;
            }
        }

        static double ConvertUOMQuantity(double qty)
        {
            return qty / unit.FromUnit * unit.ToUnit;
        }
    }
}
