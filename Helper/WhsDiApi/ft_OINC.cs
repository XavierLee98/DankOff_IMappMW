using IMAppSapMidware_NetCore.Models.SAPModels;
using System;
using System.Data;
using System.Runtime.InteropServices;

namespace IMAppSapMidware_NetCore.Helper.SQL
{
    class ft_OINC
    {
        public static string LastSAPMsg { get; set; } = string.Empty;
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

            string sapdb = Program._ErpDbName; //"SBODEMOUS2";
            string request = "Update Inventry Count";
            try
            {
                dt = ft_General.LoadData("LoadOINC_sp");
                dtDetails = ft_General.LoadDataByRequest("LoadDetails_sp", request);
                dtBin = ft_General.LoadDataByRequest("LoadBinDetails_sp", request);
                string failed_status = "ONHOLD";
                string success_status = "SUCCESS";
                string tablename = "OINC";
                string docnum = "";
                string docEntry = "";
                int cnt = 0, bin_cnt = 0, batch_cnt = 0, serial_cnt = 0, batchbin_cnt = 0;

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

                    SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        
                    SAPbobsCOM.CompanyService oCS = sap.oCom.GetCompanyService();
                    SAPbobsCOM.InventoryCountingsService oCTs = (SAPbobsCOM.InventoryCountingsService)oCS.GetBusinessService(SAPbobsCOM.ServiceTypes.InventoryCountingsService);
                    SAPbobsCOM.InventoryCounting oCT = null;// oCTs.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCounting);
                    SAPbobsCOM.InventoryCountingLine oCTl = null;
                    SAPbobsCOM.InventoryCountingBatchNumber oCTB = null;
                    SAPbobsCOM.InventoryCountingSerialNumber oCTS = null;
                    SAPbobsCOM.InventoryCountingParams oCTp = (SAPbobsCOM.InventoryCountingParams)oCTs.GetDataInterface(SAPbobsCOM.InventoryCountingsServiceDataInterfaces.icsInventoryCountingParams);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (cnt > 0)
                        {
                            if (key == dt.Rows[i]["key"].ToString()) goto details;

                            try
                            {
                                oCTs.Update(oCT);
                                //docEntry = oCTp.DocumentEntry.ToString();
                                docnum = ft_General.GetDocNum(sap.oCom, tablename, docEntry);
                                if (sap.oCom.InTransaction)
                                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                Log($" {key }\n {success_status }\n  { docnum } \n");
                                ft_General.UpdateStatus(key, success_status, "", docnum);
                            }
                            catch (Exception ex)
                            {
                                if (sap.oCom.InTransaction)
                                    sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                Log($"{key }\n {failed_status }\n { ex.Message } \n");
                                ft_General.UpdateStatus(key, failed_status, ex.Message, "");
                            }

                            cnt = 0;
                            if (oCT != null) Marshal.ReleaseComObject(oCT);
                            oCT = null;
                        }

                        if (!sap.oCom.InTransaction)
                            sap.oCom.StartTransaction();

                        docEntry = dt.Rows[i]["baseentry"].ToString();
                        oCTp.DocumentEntry = int.Parse(dt.Rows[i]["baseentry"].ToString());
                        oCT = oCTs.Get(oCTp);

                        oCT.CountDate = DateTime.Parse(dt.Rows[i]["docdate"].ToString());
                        oCT.CountTime = DateTime.Parse(dt.Rows[i]["docdate"].ToString());

                    details:

                        oCTl = oCT.InventoryCountingLines.Item(int.Parse(dt.Rows[i]["baseline"].ToString()) - 1);

                        oCTl.CountedQuantity = double.Parse(dt.Rows[i]["quantity"].ToString());

                        oCTl.Counted = SAPbobsCOM.BoYesNoEnum.tYES;


                        // add in the pur in remark 
                        // 20210403
                        var itemDetails = dt.Rows[i]["remarks"].ToString();
                        if (!string.IsNullOrWhiteSpace(itemDetails))
                        {
                            oCTl.Remarks = itemDetails;
                        }

                        DataRow[] drBin = dtBin.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'");
                        if (drBin.Length > 0 && drBin[0]["binabsentry"].ToString() != "-1") oCTl.BinEntry = int.Parse(drBin[0]["binabsentry"].ToString());

                        //DataTable dtBinBatchSerial = ft_General.LoadBinBatchSerial(dt.Rows[i]["key"].ToString(), dt.Rows[i]["itemcode"].ToString());
                        DataRow[] dr = dtDetails.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'");
                        if (dr.Length > 0)
                        {
                            if (dr[0]["batchnumber"].ToString() != "")
                            {
                                SAPbobsCOM.InventoryCountingBatchNumbers oCTBs = oCTl.InventoryCountingBatchNumbers;
                                for (int a = 0; a < oCTBs.Count; a++)
                                {
                                    oCTBs.Remove(0);
                                }
                            }
                            if (dr[0]["serialnumber"].ToString() != "")
                            {
                                SAPbobsCOM.InventoryCountingSerialNumbers oCTSs = oCTl.InventoryCountingSerialNumbers;
                                for (int a = 0; a < oCTSs.Count; a++)
                                {
                                    oCTSs.Remove(0);
                                }
                            }
                            for (int x = 0; x < dr.Length; x++)
                            {
                                if (dr[x]["batchnumber"].ToString() != "")
                                {
                                    oCTB = oCTl.InventoryCountingBatchNumbers.Add();
                                    //oCTl.TargetLine..BatchNumbers.SetCurrentLine(batch_cnt);
                                    oCTB.BatchNumber = dr[x]["batchnumber"].ToString();
                                    var batch = dr[x]["batchnumber"].ToString();
                                    oCTB.Quantity = double.Parse(decimal.Parse(dr[x]["quantity"].ToString()).ToString());
                                    var result = double.Parse(decimal.Parse(dr[x]["quantity"].ToString()).ToString());
                                    oCTB.ManufacturerSerialNumber = dr[x]["batchattr1"].ToString();

                                    if (dr[x]["batchadmissiondate"].ToString() != "")
                                        oCTB.AddmisionDate = DateTime.Parse(dr[x]["batchadmissiondate"].ToString());

                                    batch_cnt++;
                                }
                                else if (dr[x]["serialnumber"].ToString() != "")
                                {
                                    oCTS = oCTl.InventoryCountingSerialNumbers.Add();
                                    oCTS.InternalSerialNumber = dr[x]["serialnumber"].ToString();
                                    serial_cnt++;

                                    serial_cnt++;
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

                    try
                    {
                        oCTs.Update(oCT);
                        //docEntry = oCTp.DocumentEntry.ToString();
                        docnum = ft_General.GetDocNum(sap.oCom, tablename, docEntry);
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        Log($" {key }\n {success_status }\n  { docnum } \n");
                        ft_General.UpdateStatus(key, success_status, "", docnum);
                    }
                    catch (Exception ex)
                    {
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        Log($"{key }\n {failed_status }\n { ex.Message } \n");
                        ft_General.UpdateStatus(key, failed_status, ex.Message, "");
                    }

                    if (oCT != null) Marshal.ReleaseComObject(oCT);
                    oCT = null;
                }
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
                ft_General.UpdateError("OINC", ex.Message);
            }
            finally
            {
                dt = null;
                dtDetails = null;
                dtBin = null;
            }
        }
    }
}
