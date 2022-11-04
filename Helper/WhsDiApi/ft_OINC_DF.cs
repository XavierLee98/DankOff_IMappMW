using Dapper;
using IMAppSapMidware_NetCore.Helper.SQL;
using IMAppSapMidware_NetCore.Models.SAPModels;
using Microsoft.Data.SqlClient;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace IMAppSapMidware_NetCore.Helper.WhsDiApi
{
    public class ft_OINC_DF
    {
        public static string LastSAPMsg { get; set; } = string.Empty;

        static DataTable dt = null;
        static DataTable dtDetails = null;
        static DataTable dtBin = null;
        static SAPParam par;
        static SAPCompany sap;

        static void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }

        public static void Post()
        {
            string request = "Create Inventory Count";

            try
            {
                string failed_status = "ONHOLD";
                string success_status = "SUCCESS";
                string docnum = "";
                string docEntry = "";
                string tablename = "OINC";
                int retcode = 0;
                bool isOtherUOM = false;
                double convertQty = 0;

                LoadDataToDataTable(request);

                if (dt.Rows.Count > 0)
                {
                    par = SAP.GetSAPUser();
                    sap = SAP.getSAPCompany(par);

                    if (!sap.connectSAP())
                    {
                        Log($"{sap.errMsg}");
                        throw new Exception(sap.errMsg);
                    }

                    if (!sap.oCom.InTransaction)
                        sap.oCom.StartTransaction();

                    string key = dt.Rows[0]["key"].ToString();

                    CompanyService oCS = sap.oCom.GetCompanyService();
                    InventoryCountingsService oICS = (InventoryCountingsService)oCS.GetBusinessService(ServiceTypes.InventoryCountingsService);
                    InventoryCounting oIC = (InventoryCounting)oICS.GetDataInterface(InventoryCountingsServiceDataInterfaces.icsInventoryCounting);
                    InventoryCountingLines countingLines = oIC.InventoryCountingLines;
                    InventoryCountingBatchNumber batch = null;
                    InventoryCountingSerialNumber serial = null;

                    oIC.CountDate = DateTime.Parse(dt.Rows[0]["DateKey"].ToString());
                    oIC.Remarks = dt.Rows[0]["Comments"].ToString();

                    if (!string.IsNullOrEmpty(dt.Rows[0]["UserListStr"].ToString()))
                    {
                        oIC.UserFields.Item("U_UserList").Value = dt.Rows[0]["UserListStr"].ToString();
                    }

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        InventoryCountingLine countingLine = countingLines.Add();

                        countingLine.ItemCode = dt.Rows[i]["ItemCode"].ToString();
                        countingLine.UoMCode = dt.Rows[i]["UomCode"].ToString();
                        if (dt.Rows[i]["UomCode"].ToString() == "Manual")
                        {
                            countingLine.CountedQuantity = double.Parse(dt.Rows[i]["quantity"].ToString());
                        }
                        else
                        {
                            countingLine.UoMCountedQuantity = double.Parse(dt.Rows[i]["quantity"].ToString());
                            isOtherUOM = true;
                        }
                        var test = dt.Rows[i]["UomCode"].ToString();

                        countingLine.WarehouseCode = dt.Rows[i]["whscode"].ToString();
                        countingLine.BinEntry = int.Parse(dt.Rows[i]["BinAbsEntry"].ToString());
                        countingLine.Counted = SAPbobsCOM.BoYesNoEnum.tYES;

                        DataRow[] dr = dtDetails.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "' and BcdCode='" + dt.Rows[i]["BcdCode"].ToString() + "'");

                        if (dr.Length > 0)
                        {
                            for (int x = 0; x < dr.Length; x++)
                            {
                                if (dr[x]["batchnumber"].ToString() != "")
                                {
                                    batch = countingLine.InventoryCountingBatchNumbers.Add();
                                    batch.BatchNumber = dr[x]["batchnumber"].ToString();
                                    if(!isOtherUOM)
                                        batch.Quantity = double.Parse(dr[x]["Quantity"].ToString());
                                    else
                                    {
                                        batch.Quantity = ConvertToInventoryUOM(dt.Rows[i]["UomCode"].ToString(), double.Parse(dr[x]["Quantity"].ToString()));
                                    }
                                }
                                else if (dr[x]["serialnumber"].ToString() != "")
                                {
                                    serial = countingLine.InventoryCountingSerialNumbers.Add();
                                    serial.InternalSerialNumber = dr[x]["serialnumber"].ToString();
                                    serial.Quantity = double.Parse(dr[x]["Quantity"].ToString());
                                }
                            }
                        }
                    }

                    try
                    {
                        SAPbobsCOM.InventoryCountingParams oICP = oICS.Add(oIC);
                        docnum = oICP.DocumentEntry.ToString();

                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                        UpdateDraft(key);
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
                }
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
                ft_General.UpdateError("OINC", ex.Message);
            }
        }

        static void LoadDataToDataTable(string request)
        {
            dt = ft_General.LoadData("LoadInventoryCount_sp");
            dtDetails = ft_General.LoadDataByRequest("LoadDetails_sp", request);
            dtBin = ft_General.LoadDataByRequest("LoadBinDetails_sp", request);
        }

        static void UpdateDraft(string PostGuid)
        {
            try
            {
                var conn = new System.Data.SqlClient.SqlConnection(Program._DbMidwareConnStr);
                string updateQuery = "UPDATE zmwInventoryCountHead SET DocStatus = 'Posted' WHERE PostGuid = @PostGuid";
                var result = conn.Execute(updateQuery, new { PostGuid });
            }
            catch (Exception e)
            {
                Log(e.ToString());
            }
        }
        static double ConvertToInventoryUOM(string FromUomCode, double qty)
        {
            try
            {
                var conn = new System.Data.SqlClient.SqlConnection(Program._DbErpConnStr);
                string query = $"SELECT T1.AltQty [FromUnit], T1.BaseQty [ToUnit] FROM OUOM T0 " +
                               $"INNER JOIN UGP1 T1 on T1.UomEntry = T0.UomEntry " +
                               $"WHERE T0.UomCode = @UomCode";

                var convertUnit = conn.Query<UOMConvert>(query, new { UomCode = FromUomCode }).FirstOrDefault();

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

