using Dapper;
using IMAppSapMidware_NetCore.Helper.SQL;
using IMAppSapMidware_NetCore.Models.SAPModels;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace IMAppSapMidware_NetCore.Helper.WhsDiApi
{
    public class ft_OPKL
    {
        public static string LastSAPMsg { get; set; } = string.Empty;

        static string currentKey = string.Empty;
        static string currentStatus = string.Empty;
        static string CurrentDocNum = string.Empty;
        public static string Erp_DBConnStr { get; set; } = string.Empty;


        static DataTable dt = null;
        static DataTable dtDetails = null;
        static DataTable dtBin = null;
        static SAPParam par;
        static SAPCompany sap;
        static PickLists oPickLists = null;
        static PickLists_Lines oPickLists_Lines = null;
        static bool isOtherUOM = false;
        static UOMConvert unit = null;

        static void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }

        public async static void Post()
        {
            string request = "Update Pick List";
            try
            {
                string failed_status = "ONHOLD";
                string success_status = "SUCCESS";
                int cnt = 0, bin_cnt = 0, batch_cnt = 0, serial_cnt = 0, batchbin_cnt = 0;
                int retcode = 0;

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
                    string key = dt.Rows[0]["key"].ToString();
                    currentKey = key;
                    currentStatus = failed_status;

                    if (!sap.oCom.InTransaction)
                        sap.oCom.StartTransaction();

                    oPickLists = (SAPbobsCOM.PickLists)sap.oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);
                    oPickLists.GetByKey(int.Parse(dt.Rows[0]["sapDocNumber"].ToString()));

                    if (!CheckIsEcommerce(int.Parse(dt.Rows[0]["sapDocNumber"].ToString())))
                    {
                        if (dt.Rows[0]["CartonType"].ToString() != "")
                             oPickLists.UserFields.Fields.Item("U_CartonSizeType").Value = dt.Rows[0]["CartonType"].ToString();
                        if (dt.Rows[0]["CartonSize"].ToString() != "")
                            oPickLists.UserFields.Fields.Item("U_CartonSize").Value = dt.Rows[0]["CartonSize"].ToString();
                        if (dt.Rows[0]["AirwayBill"].ToString() != "")
                            oPickLists.UserFields.Fields.Item("U_AirwayBill").Value = dt.Rows[0]["AirwayBill"].ToString();
                        if (dt.Rows[0]["TotalWeight"].ToString() != "")
                            oPickLists.UserFields.Fields.Item("U_TotalWeight").Value = double.Parse(dt.Rows[0]["TotalWeight"].ToString());

                        oPickLists.UserFields.Fields.Item("U_PickDate").Value = DateTime.Now;
                    }

                    var result = dt.Rows[0]["IsCompletePick"].ToString();
                    if (dt.Rows[0]["IsCompletePick"].ToString() != "" && bool.Parse(dt.Rows[0]["IsCompletePick"].ToString()))
                    {
                        oPickLists.UserFields.Fields.Item("U_IsCompletePicked").Value = "Y";
                        oPickLists.UserFields.Fields.Item("U_PickDate").Value = DateTime.Now;
                    }

                    CurrentDocNum = dt.Rows[0]["sapDocNumber"].ToString();
                    oPickLists_Lines = oPickLists.Lines;

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        isOtherUOM = false;
                        unit = null;
                        oPickLists_Lines.SetCurrentLine(int.Parse(dt.Rows[i]["SourceLineNum"].ToString()));
                        if(dt.Rows[i]["UomCode"].ToString() != "Manual")
                        {
                            isOtherUOM = true;
                            unit = await GetUOMUnit(dt.Rows[i]["UomCode"].ToString());
                            if (unit == null) throw new Exception("UOM Unit is null.");

                            oPickLists_Lines.PickedQuantity = ConvertUOMQuantity(double.Parse(dt.Rows[i]["quantity"].ToString()));
                        }
                        else
                            oPickLists_Lines.PickedQuantity = double.Parse(dt.Rows[i]["quantity"].ToString());

                        DataRow[] dr = dtDetails.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and SourceLineNum='" + dt.Rows[i]["SourceLineNum"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'");

                        if (dr.Length > 0)
                        {
                            for (int x = 0; x < dr.Length; x++)
                            {
                                if (dr[x]["batchnumber"].ToString() != "")
                                {
                                    PerformBatchTransaction(dr[x], dt.Rows[i]["key"].ToString());
                                    batch_cnt++;
                                }
                                else if (dr[x]["serialnumber"].ToString() != "")
                                {
                                    PerformSerialTransaction(dr[x], dt.Rows[i]["key"].ToString(), dt.Rows[i]["itemcode"].ToString());
                                    serial_cnt++;
                                }
                                else
                                {
                                    DataRow[] drBin = dtBin.Select("guid='" + dt.Rows[i]["key"].ToString() + "' and SourceLineNum ='" + dt.Rows[i]["SourceLineNum"].ToString() + "' and itemcode='" + dt.Rows[i]["itemcode"].ToString() + "'");
                                    if (drBin.Length > 0)
                                    {
                                        PerformNormalItemTransaction(drBin);
                                    }
                                }
                            }
                        }
                    }

                    retcode = oPickLists.Update();

                    if (retcode == 0)
                    {
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                        Log($"{key }\n {success_status }\n  { CurrentDocNum } \n");
                        ft_General.UpdateStatus(key, success_status, "", CurrentDocNum);
                    }
                    else
                    {
                        if (sap.oCom.InTransaction)
                            sap.oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        string message = sap.oCom.GetLastErrorDescription().ToString().Replace("'", "");
                        Log($"{key }\n {failed_status }\n { message } \n");
                        ft_General.UpdateStatus(key, failed_status, message, CurrentDocNum);
                        //UpdateIsCompletedPickedToNo(int.Parse(CurrentDocNum));
                    }

                    if (oPickLists != null) Marshal.ReleaseComObject(oPickLists);
                    oPickLists = null;

                }
            }
            catch (Exception ex)
            {
                Log($"{ ex.Message } \n");
                ft_General.UpdateError("OPKL", ex.Message);

                Log($"{currentKey }\n {currentStatus }\n { ex.Message } \n");
                ft_General.UpdateStatus(currentKey, currentStatus, ex.Message, CurrentDocNum);
                //UpdateIsCompletedPickedToNo(int.Parse(CurrentDocNum));
            }
            finally
            {
                dt = null;
                dtDetails = null;
                dtBin = null;
            }
        }

        static void LoadDataToDataTable(string request)
        {
            dt = ft_General.LoadData("LoadOPKL_sp");
            dt.DefaultView.Sort = "key, SourceLineNum";
            dt = dt.DefaultView.ToTable();

            dtDetails = ft_General.LoadDataByRequest("LoadDetails_sp", request);
            dtBin = ft_General.LoadDataByRequest("LoadBinDetails_sp", request);
        }

        static void PerformSerialTransaction(DataRow row, string key, string itemCode)
        {
            int serial_cnt = 0;
            bool found = false;

            for (int x = 0; x < oPickLists_Lines.SerialNumbers.Count; x++)
            {
                oPickLists_Lines.SerialNumbers.SetCurrentLine(x);
                if (oPickLists_Lines.SerialNumbers.InternalSerialNumber == row["serialnumber"].ToString())
                {
                    oPickLists_Lines.SerialNumbers.Quantity = 1;
                    oPickLists_Lines.SerialNumbers.BaseLineNumber = int.Parse(row["SourceLineNum"].ToString());
                    serial_cnt = x;
                    found = true;
                    break;
                }
            }

            DataRow[] drBin = dtBin.Select("guid='" + key + "' and itemcode='" + itemCode + "' and SourceLineNum ='" + row["SourceLineNum"].ToString() +
                                         "' and serialnumber ='" + row["serialnumber"].ToString() + "'");
            if (!found)
            {
                if (oPickLists_Lines.SerialNumbers.Count == 0)
                    oPickLists_Lines.SerialNumbers.Add();
                else
                {
                    oPickLists_Lines.SerialNumbers.SetCurrentLine(oPickLists_Lines.SerialNumbers.Count - 1);
                    if (!string.IsNullOrEmpty(oPickLists_Lines.SerialNumbers.InternalSerialNumber))
                        oPickLists_Lines.SerialNumbers.Add();
                }
                oPickLists_Lines.SerialNumbers.InternalSerialNumber = row["serialnumber"].ToString();
                oPickLists_Lines.SerialNumbers.Quantity = 1;
                oPickLists_Lines.SerialNumbers.BaseLineNumber = int.Parse(row["SourceLineNum"].ToString());
                serial_cnt = oPickLists_Lines.SerialNumbers.Count - 1;

                if (drBin.Length > 0)
                {
                    for (int y = 0; y < drBin.Length; y++)
                    {
                        if (oPickLists_Lines.BinAllocations.Count == 0)
                            oPickLists_Lines.BinAllocations.Add();
                        else
                        {
                            oPickLists_Lines.BinAllocations.SetCurrentLine(oPickLists_Lines.BinAllocations.Count - 1);
                            if (oPickLists_Lines.BinAllocations.BinAbsEntry > 0)
                                oPickLists_Lines.BinAllocations.Add();
                        }
                        oPickLists_Lines.BinAllocations.BinAbsEntry = int.Parse(drBin[y]["binabsentry"].ToString());
                        oPickLists_Lines.BinAllocations.Quantity = double.Parse(drBin[y]["quantity"].ToString());
                        oPickLists_Lines.BinAllocations.BaseLineNumber = int.Parse(row["SourceLineNum"].ToString());
                        oPickLists_Lines.BinAllocations.SerialAndBatchNumbersBaseLine = serial_cnt;
                    }
                }
            }
        }

        static void PerformBatchTransaction(DataRow row, string key)
        {
            int batch_cnt = 0;
            bool found = false;
            for (int x = 0; x < oPickLists_Lines.BatchNumbers.Count; x++)
            {
                oPickLists_Lines.BatchNumbers.SetCurrentLine(x);
                if (oPickLists_Lines.BatchNumbers.BatchNumber == row["batchnumber"].ToString())
                {
                    if (isOtherUOM)
                    {
                        oPickLists_Lines.BatchNumbers.Quantity = ConvertUOMQuantity(double.Parse(row["Quantity"].ToString()));
                    }
                    else
                    {
                        oPickLists_Lines.BatchNumbers.Quantity = double.Parse(row["Quantity"].ToString());
                    }
                    oPickLists_Lines.BatchNumbers.BaseLineNumber = int.Parse(row["SourceLineNum"].ToString());
                    batch_cnt = x;
                    found = true;
                    break;
                }
            }

            DataRow[] drBin = dtBin.Select("guid='" + key + "' and itemcode='" + row["itemcode"].ToString() + "' and SourceLineNum ='" + row["SourceLineNum"].ToString() +
                "' and Batchnumber ='" + row["batchnumber"].ToString() + "'");

            if (!found)
            {
                if (oPickLists_Lines.BatchNumbers.Count == 0)
                    oPickLists_Lines.BatchNumbers.Add();
                else
                {
                    oPickLists_Lines.BatchNumbers.SetCurrentLine(oPickLists_Lines.BatchNumbers.Count - 1);
                    if (!string.IsNullOrEmpty(oPickLists_Lines.BatchNumbers.BatchNumber))
                        oPickLists_Lines.BatchNumbers.Add();
                }

                oPickLists_Lines.BatchNumbers.BatchNumber = row["batchnumber"].ToString();
                if (isOtherUOM)
                {
                    oPickLists_Lines.BatchNumbers.Quantity = ConvertUOMQuantity(double.Parse(row["Quantity"].ToString()));
                }
                else
                {
                    oPickLists_Lines.BatchNumbers.Quantity = double.Parse(row["Quantity"].ToString());
                }
                //oPickLists_Lines.BatchNumbers.Quantity = double.Parse(row["Quantity"].ToString());
                oPickLists_Lines.BatchNumbers.BaseLineNumber = int.Parse(row["SourceLineNum"].ToString());
                batch_cnt = oPickLists_Lines.BatchNumbers.Count - 1;

                if (drBin.Length > 0)
                {
                    for (int y = 0; y < drBin.Length; y++)
                    {
                        if (oPickLists_Lines.BinAllocations.Count == 0)
                            oPickLists_Lines.BinAllocations.Add();
                        else
                        {
                            oPickLists_Lines.BinAllocations.SetCurrentLine(oPickLists_Lines.BinAllocations.Count - 1);
                            if (oPickLists_Lines.BinAllocations.BinAbsEntry > 0)
                                oPickLists_Lines.BinAllocations.Add();
                        }
                        oPickLists_Lines.BinAllocations.BinAbsEntry = int.Parse(drBin[y]["binabsentry"].ToString());

                        if (isOtherUOM)
                        {
                            oPickLists_Lines.BinAllocations.Quantity = ConvertUOMQuantity(double.Parse(drBin[y]["quantity"].ToString()));
                        }
                        else
                        {
                            oPickLists_Lines.BinAllocations.Quantity = double.Parse(drBin[y]["quantity"].ToString());
                        }
                        //oPickLists_Lines.BinAllocations.Quantity = double.Parse(drBin[y]["quantity"].ToString());
                        oPickLists_Lines.BinAllocations.BaseLineNumber = int.Parse(row["SourceLineNum"].ToString());
                        oPickLists_Lines.BinAllocations.SerialAndBatchNumbersBaseLine = batch_cnt;

                    }
                }
            }
            else
            {
                if (drBin.Length > 0)
                {
                    for (int y = 0; y < drBin.Length; y++)
                    {
                        found = false;
                        for (int x = 0; x < oPickLists_Lines.BinAllocations.Count; x++)
                        {
                            oPickLists_Lines.BinAllocations.SetCurrentLine(x);
                            if (oPickLists_Lines.BinAllocations.BinAbsEntry == int.Parse(drBin[y]["binabsentry"].ToString()) 
                                && oPickLists_Lines.BinAllocations.BaseLineNumber == int.Parse(row["SourceLineNum"].ToString())
                                && oPickLists_Lines.BinAllocations.SerialAndBatchNumbersBaseLine == batch_cnt)
                            {
                                found = true;
                                break;
                            }
                        }

                        if (!found)
                        {
                            if (oPickLists_Lines.BinAllocations.Count == 0)
                                oPickLists_Lines.BinAllocations.Add();
                            else
                            {
                                oPickLists_Lines.BinAllocations.SetCurrentLine(oPickLists_Lines.BinAllocations.Count - 1);
                                if (oPickLists_Lines.BinAllocations.BinAbsEntry > 0)
                                    oPickLists_Lines.BinAllocations.Add();
                            }
                        }
                        oPickLists_Lines.BinAllocations.BinAbsEntry = int.Parse(drBin[y]["binabsentry"].ToString());

                        if (isOtherUOM)
                        {
                            oPickLists_Lines.BinAllocations.Quantity = ConvertUOMQuantity(double.Parse(drBin[y]["quantity"].ToString()));
                        }
                        else
                        {
                            oPickLists_Lines.BinAllocations.Quantity = double.Parse(drBin[y]["quantity"].ToString());
                        }
                        oPickLists_Lines.BinAllocations.BaseLineNumber = int.Parse(row["SourceLineNum"].ToString());
                        oPickLists_Lines.BinAllocations.SerialAndBatchNumbersBaseLine = batch_cnt;
                    }
                }

            }
        }

        static void PerformNormalItemTransaction(DataRow[] drBin)
        {
            int bin_cnt = 0;
            bool found = false;

            for (int y = 0; y < drBin.Length; y++)
            {
                for (int x = 0; x < oPickLists_Lines.BinAllocations.Count; x++)
                {
                    oPickLists_Lines.BinAllocations.SetCurrentLine(x);
                    if (oPickLists_Lines.BinAllocations.BinAbsEntry == int.Parse(drBin[y]["binabsentry"].ToString()))
                    {
                        if (isOtherUOM)
                        {
                            oPickLists_Lines.BinAllocations.Quantity = ConvertUOMQuantity(double.Parse(drBin[y]["quantity"].ToString()));
                        }
                        else
                        {
                            oPickLists_Lines.BinAllocations.Quantity = double.Parse(drBin[y]["quantity"].ToString());
                        }
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    if (oPickLists_Lines.BinAllocations.Count == 0)
                        oPickLists_Lines.BinAllocations.Add();
                    else
                    {
                        oPickLists_Lines.BinAllocations.SetCurrentLine(oPickLists_Lines.BinAllocations.Count - 1);
                        if (oPickLists_Lines.BinAllocations.BinAbsEntry > 0)
                            oPickLists_Lines.BinAllocations.Add();
                    }

                    oPickLists_Lines.BinAllocations.BinAbsEntry = int.Parse(drBin[y]["binabsentry"].ToString());
                    if (isOtherUOM)
                    {
                        oPickLists_Lines.BinAllocations.Quantity = ConvertUOMQuantity(double.Parse(drBin[y]["quantity"].ToString()));
                    }
                    else
                    {
                        oPickLists_Lines.BinAllocations.Quantity = double.Parse(drBin[y]["quantity"].ToString());
                    }
                    oPickLists_Lines.BinAllocations.BaseLineNumber = int.Parse(drBin[y]["SourceLineNum"].ToString());
                }
            }
        }

        static async Task<UOMConvert> GetUOMUnit(string FromUomCode)
        {
            try
            {
                var conn = new SqlConnection(Program._DbErpConnStr);
                string query = $"SELECT T1.AltQty [FromUnit], T1.BaseQty [ToUnit] FROM OUOM T0 " +
                               $"INNER JOIN UGP1 T1 on T1.UomEntry = T0.UomEntry " +
                               $"WHERE T0.UomCode = @UomCode";

                return  conn.Query<UOMConvert>(query, new { UomCode = FromUomCode }).FirstOrDefault();
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

        static bool CheckIsEcommerce(int PickNo)
        {
            var conn = new SqlConnection(Program._DbErpConnStr);
            
            var result = conn.Query<string>("zwa_IMApp_PickList_spCheckIsEcommerceWithPickNo", 
                new { AbsEntry = PickNo },
                commandType:CommandType.StoredProcedure)
                .FirstOrDefault();

            return  result == "Y" ? true : false;
        }

        static void UpdateIsCompletedPickedToNo(int PickNo)
        {
            var conn = new SqlConnection(Program._DbErpConnStr);
            
            var result = conn.Query<string>("zwa_IMApp_PickList_spUpdateIsCompletedPickedToNo", 
                new { AbsEntry = PickNo },
                commandType: CommandType.StoredProcedure
                ).FirstOrDefault();

            return;
        }



    }
}

