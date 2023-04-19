using Dapper;
using IMAppSapMidware_NetCore.Models.SAPModels;
using SAPbobsCOM;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace IMAppSapMidware_NetCore.Helper.SQL
{
    class ft_OIGE
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
            DataTable dtDetails = null;
            DataTable dtBin = null;
            DataTable dtAttachments = null;

            string request = "Create GI";
            string sapdb = "";
            try
            {
                dt = ft_General.LoadData("LoadOIGE_sp");
                dtDetails = ft_General.LoadDataByRequest("LoadDetails_sp", request);
                dtBin = ft_General.LoadDataByRequest("LoadBinDetails_sp", request);
                string failed_status = "ONHOLD";
                string success_status = "SUCCESS";
                string tablename = "OIGE";
                string docnum = "";
                string docEntry = "";
                int cnt = 0, bin_cnt = 0, batch_cnt = 0, serial_cnt = 0, batchbin_cnt = 0;
                int retcode = 0;
                int iAttEntry = -1;
                int UomEntry = -1;
                double convertedQuantity = -1;

                if (dt.Rows.Count > 0)
                {
                    SAPParam par = SAP.GetSAPUser();
                    SAPCompany sap = SAP.getSAPCompany(par);

                    if (!sap.connectSAP())
                    {
                        Log($"{sap.errMsg}");
                        throw new Exception(sap.errMsg);
                    }

                    DateTime docdate = DateTime.Parse(dt.Rows[0]["docdate"].ToString());
                    string key = dt.Rows[0]["key"].ToString();

                    currentKey = key;
                    currentStatus = failed_status;

                    //SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Documents oDoc = null;// (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                    SAPbobsCOM.Attachments2 oATT = null;

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

                        oDoc = (SAPbobsCOM.Documents)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
                        oDoc.DocObjectCodeEx = ((int)SAPbobsCOM.BoObjectTypes.oInventoryGenExit).ToString();
                        oDoc.DocDate = DateTime.Parse(dt.Rows[i]["docdate"].ToString());
                        oDoc.TaxDate = DateTime.Parse(dt.Rows[i]["docdate"].ToString());

                        if (dt.Rows[i]["series"].ToString() != "")
                            oDoc.Series = int.Parse(dt.Rows[i]["series"].ToString());
                        if (dt.Rows[i]["ref2"].ToString() != "")
                            oDoc.Reference2 = dt.Rows[i]["ref2"].ToString();
                        if (dt.Rows[i]["comments"].ToString() != "")
                            oDoc.Comments = dt.Rows[i]["comments"].ToString();
                        if (dt.Rows[i]["jrnlmemo"].ToString() != "")
                            oDoc.JournalMemo = dt.Rows[i]["jrnlmemo"].ToString();

                        //if (dt.Rows[i]["numatcard"].ToString() != "")
                        //    oDoc.NumAtCard = dt.Rows[i]["numatcard"].ToString();

                        //#region Attachments
                        //dtAttachments = ft_General.LoadDataByGuid("LoadAttachments_sp", dt.Rows[i]["key"].ToString());

                        //if (dtAttachments.Rows.Count > 0)
                        //{
                        //    oATT = (Attachments2)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);
                        //    for (int x = 0; x < dtAttachments.Rows.Count; x++)
                        //    {
                        //        string filePath = $"{dtAttachments.Rows[x]["serverSavedPath"].ToString()}";

                        //        if (File.Exists(filePath))
                        //        {
                        //            oATT.Lines.Add();
                        //            oATT.Lines.FileName = Path.GetFileNameWithoutExtension(filePath);
                        //            oATT.Lines.FileExtension = Path.GetExtension(filePath).Substring(1);
                        //            oATT.Lines.SourcePath = Path.GetDirectoryName(filePath);
                        //            oATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;
                        //        }
                        //    }
                        //    retcode = oATT.Add();
                        //    if (retcode == 0)
                        //    {
                        //        iAttEntry = int.Parse(sap.oCom.GetNewObjectKey());
                        //        //Assign the attachment to the GR object (GR is my SAPbobsCOM.Documents object)
                        //    }
                        //    else
                        //    {
                        //        string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
                        //        ft_General.UpdateStatus(key, failed_status, message, "");
                        //        continue;
                        //    }
                        //    if (oATT != null) Marshal.ReleaseComObject(oATT);
                        //    oATT = null;
                        //}

                        //if (iAttEntry != -1) oDoc.AttachmentEntry = iAttEntry;
                        //#endregion

                        details:
                        oDoc.Lines.ItemCode = dt.Rows[i]["itemcode"].ToString();
                        oDoc.Lines.UseBaseUnits = SAPbobsCOM.BoYesNoEnum.tNO;
                        oDoc.Lines.UoMEntry = int.Parse(dt.Rows[i]["UomEntry"].ToString());
                        UomEntry = int.Parse(dt.Rows[i]["UomEntry"].ToString());
                        oDoc.Lines.Quantity = double.Parse(dt.Rows[i]["quantity"].ToString());
                        oDoc.Lines.WarehouseCode = dt.Rows[i]["whscode"].ToString();

                        var itemDetails = dt.Rows[i]["remarks"].ToString();
                        if (!string.IsNullOrWhiteSpace(itemDetails))
                        {
                            oDoc.Lines.ItemDetails = itemDetails;
                        }

                        DataRow[] dr = dtDetails.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'" +
                                         " and lineguid='" + dt.Rows[i]["lineguid"].ToString() + "'");

                        //DataRow[] dr = dtDetails.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'");
                        if (dr.Length > 0)
                        {
                            for (int x = 0; x < dr.Length; x++)
                            {
                                if (dr[x]["batchnumber"].ToString() != "")
                                {
                                    if (batch_cnt > 0) oDoc.Lines.BatchNumbers.Add();
                                    //oDoc.Lines.BatchNumbers.SetCurrentLine(batch_cnt);
                                    oDoc.Lines.BatchNumbers.BatchNumber = dr[x]["batchnumber"].ToString();

                                    convertedQuantity = ConvertToInventoryUOM(dt.Rows[i]["itemcode"].ToString(), UomEntry, double.Parse(dr[x]["quantity"].ToString()));
                                    if (convertedQuantity < 0) throw new Exception(LastSAPMsg);

                                    oDoc.Lines.BatchNumbers.Quantity = convertedQuantity;

                                    oDoc.Lines.BatchNumbers.ManufacturerSerialNumber = dr[x]["batchattr1"].ToString();

                                    if (dr[x]["admissiondate"].ToString() != "")
                                        oDoc.Lines.BatchNumbers.AddmisionDate = DateTime.Parse(dr[x]["admissiondate"].ToString());

                                    //DataRow[] drBin = dtBin.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() +
                                    //    "' and Batchnumber ='" + dr[x]["batchnumber"].ToString() + "'");

                                    DataRow[] drBin = dtBin.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'" +
                                        " and Batchnumber ='" + dr[x]["batchnumber"].ToString() + "'" +
                                        " and lineguid='" + dt.Rows[i]["lineguid"].ToString() + "'");


                                    if (drBin.Length > 0)
                                    {
                                        for (int y = 0; y < drBin.Length; y++)
                                        {
                                            if (batchbin_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                           // oDoc.Lines.BinAllocations.SetCurrentLine(batchbin_cnt);
                                            oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBin[y]["binabsentry"].ToString());

                                            convertedQuantity = ConvertToInventoryUOM(dt.Rows[i]["itemcode"].ToString(), UomEntry, double.Parse(drBin[y]["quantity"].ToString()));
                                            if (convertedQuantity < 0) throw new Exception(LastSAPMsg);

                                            oDoc.Lines.BinAllocations.Quantity = convertedQuantity;

                                            oDoc.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = batch_cnt;
                                            batchbin_cnt++;
                                        }
                                    }

                                    batch_cnt++;
                                }
                                else if (dr[x]["serialnumber"].ToString() != "")
                                {
                                    if (serial_cnt > 0) oDoc.Lines.SerialNumbers.Add();
                                    //oDoc.Lines.SerialNumbers.SetCurrentLine(serial_cnt);
                                    oDoc.Lines.SerialNumbers.InternalSerialNumber = dr[x]["serialnumber"].ToString();
                                    DataRow[] drBin = dtBin.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() +
                                        "' and serialnumber ='" + dr[x]["serialnumber"].ToString() + "'");

                                    if (drBin.Length > 0)
                                    {
                                        if (serial_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                        //oDoc.Lines.BinAllocations.SetCurrentLine(serial_cnt);
                                        oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBin[0]["binabsentry"].ToString());
                                        oDoc.Lines.BinAllocations.Quantity = double.Parse(drBin[0]["quantity"].ToString());
                                        oDoc.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = serial_cnt;
                                    }

                                    serial_cnt++;
                                }
                                else
                                {
                                    DataRow[] drBin = dtBin.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'" +
                                                                   " and lineguid='" + dt.Rows[i]["lineguid"].ToString() + "'");

                                    if (drBin.Length > 0)
                                    {
                                        for (int y = 0; y < drBin.Length; y++)
                                        {
                                            if (bin_cnt > 0) oDoc.Lines.BinAllocations.Add();
                                            oDoc.Lines.BinAllocations.SetCurrentLine(bin_cnt);
                                            oDoc.Lines.BinAllocations.BinAbsEntry = int.Parse(drBin[y]["binabsentry"].ToString());

                                            convertedQuantity = ConvertToInventoryUOM(dt.Rows[i]["itemcode"].ToString(), UomEntry, double.Parse(drBin[y]["quantity"].ToString()));
                                            if (convertedQuantity < 0) throw new Exception(LastSAPMsg);
                                            oDoc.Lines.BinAllocations.Quantity = convertedQuantity;
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


                        docdate = DateTime.Parse(dt.Rows[i]["docdate"].ToString());
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
                        //docnum = ft_General.GetDraftDocNum(sap.oCom, docEntry, (int)SAPbobsCOM.BoObjectTypes.oInventoryGenExit);

                        CurrentDocNum = docEntry;

                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        Log($"{key }\n {success_status }\n  { docEntry } \n");
                        ft_General.UpdateStatus(key, success_status, "", docEntry);
                    }

                    if (oDoc != null) Marshal.ReleaseComObject(oDoc);
                    oDoc = null;
                }
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
                ft_General.UpdateError("OIGE", ex.Message);

                Log($"{currentKey }\n {currentStatus }\n { ex.Message } \n");
                ft_General.UpdateStatus(currentKey, currentStatus, ex.Message, CurrentDocNum);
            }
            finally
            {
                dt = null;
                dtDetails = null;
                dtBin = null;
                dtAttachments = null;
            }
        }

        static double ConvertToInventoryUOM(string itemcode ,int FromUomEntry, double qty)
        {
            try
            {
                var conn = new SqlConnection(Program._DbErpConnStr);
                string query = $"SELECT * FROM IMAPP_Item_ConvertToInventoryUOM (@itemcode, null, @uomentry)";

                //string query = $"SELECT T1.AltQty [FromUnit], T1.BaseQty [ToUnit] FROM OUOM T0 " +
                //               $"INNER JOIN UGP1 T1 on T1.UomEntry = T0.UomEntry " +
                //               $"WHERE T0.UomEntry = @UomEntry";

                var convertUnit = conn.Query<UOMConvert>(query, new { itemcode = itemcode, uomentry = FromUomEntry }).FirstOrDefault();

                //var convertUnit = conn.Query<UOMConvert>(query, new { UomEntry = FromUomEntry }).FirstOrDefault();

                var covertedQty = qty / convertUnit.FromUnit * convertUnit.ToUnit;

                return covertedQty;
            }
            catch (Exception e)
            {
                LastSAPMsg = e.ToString();
                return -1;
            }
        }
    }
}
