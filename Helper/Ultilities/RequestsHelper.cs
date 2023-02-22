using Dapper;
using IMAppSapMidware_NetCore.Helper.SQL;
using IMAppSapMidware_NetCore.Helper.WhsDiApi;
using IMAppSapMidware_NetCore.Models;
using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace IMAppSapMidware_NetCore.Helper
{
    public class RequestsHelper
    {
        public string LasteErrorMessage { get; set; } = string.Empty;
        public string DbConnectString_Midware { get; set; } = string.Empty;
        public string DbConnectString_Erp { get; set; } = string.Empty;
        public bool IsBusy { get; set; } = false;

        void Log(string message)
        {
            Program.FilLogger?.Log(message);
        }

        public void ExecuteRequest()
        {
            try
            {
                if (IsBusy) return;
                IsBusy = true; 

                var sql = @"SELECT * FROM zmwRequest 
                            WHERE Status = 'ONHOLD' and tried < 3";

                var requests = new List<Request>();

                using var conn = new SqlConnection(DbConnectString_Midware);
                requests = conn.Query<Request>(sql).ToList();

                if (requests == null) return;
                if (requests.Count == 0) return;

                //var erpProperty = GetProperty();
                //if (erpProperty == null)
                //{
                //    Log("Midware database [ft_SAPSettings] found empty, please configure, and try again. ");
                //    return;
                //}

                HandlerRequests(requests);
                IsBusy = false; 
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{ e.StackTrace}");
            }
        }

        //ErpPropertyHelper GetProperty()
        //{
        //    try
        //    {
        //        var sql = "SELECT * FROM ft_SAPSettings";
        //        using var conn = new SqlConnection(DbConnectString_Midware);
        //        return conn.Query<ErpPropertyHelper>(sql).FirstOrDefault();
        //    }
        //    catch (Exception e)
        //    {
        //        Log($"{e.Message}\n{ e.StackTrace}");
        //        return null;
        //    }
        //}

        void HandlerRequests(List<Request> Reqiuests)
        {
            try
            {
                Reqiuests.ForEach(request =>
                {
                    switch (request.request)
                    {
                        case "Create GRPO":
                            {
                                try
                                {
                                    //Create GRPO
                                    ft_OPDN.Post();
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("OPDN", ex.Message);
                                }
                                break;
                            }
                        case "Create DO":
                            {
                                try
                                {
                                    //Create DO
                                    ft_ODLN.Post();
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("ODLN", ex.Message);
                                }
                                break;
                            }
                        case "Create GR":
                            {
                                try
                                {
                                    //Create GR
                                    ft_OIGN.Post();
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("OIGN", ex.Message);
                                }
                                break;
                            }
                        case "Create GI":
                            {
                                try
                                {
                                    //Create GI
                                    ft_OIGE.Post();
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("OIGE", ex.Message);
                                }
                                break;
                            }
                        case "Create Inventory Request":
                            {
                                try
                                {
                                    //Create Inventory Request
                                    //LoadOWTQ_sp
                                    ft_OWTQ.Post(); // not item detals line to put remarks // 20210403
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("OWTQ", ex.Message);
                                }
                                break;
                            }
                        case "Create Transfer1":
                            {
                                try
                                {
                                    //Create Transfer1
                                    ft_OWTR.Post(); // not item detals line to put remarks // 20210403
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("OWTR", ex.Message);
                                }
                                break;
                            }
                        case "Create DF Transfer":
                            {
                                try
                                {
                                    //Create Transfer1
                                    ft_OWTR_DF.Post(); // not item detals line to put remarks // 20210403
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("OWTR", ex.Message);
                                }
                                break;
                            }
                        //case "Update Inventry Count":
                        //    {
                        //        try
                        //        {
                        //            //Update Inventry Count
                        //            ft_OINC.Post();
                        //        }
                        //        catch (Exception ex)
                        //        {
                        //            Log($"{ ex.Message } \n");
                        //            ft_General.UpdateError("OINC", ex.Message);
                        //        }
                        //        break;
                        //    }
                        case "Create Inventory Count":
                            {
                                try
                                {
                                    ft_OINC_DF.Post();
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("OINC", ex.Message);
                                }
                                break;
                            }
                        case "Create Return":
                            {
                                try
                                {
                                    //Create Return
                                    ft_ORDN.Post();
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("ORDN", ex.Message);
                                }
                                break;
                            }
                        case "Create Return Request":
                            {
                                try
                                {
                                    //LoadORRR_sp
                                    ft_ORRR.Post();
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("ORRR", ex.Message);
                                }
                                break;
                            }
                        case "Create Goods Return":
                            {
                                try
                                {
                                    //Create Goods Return
                                    ft_ORPD.Post();
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("ORPD", ex.Message);
                                }
                                break;
                            }
                        case "Create Goods Return Request":
                            {
                                try
                                {
                                    //LoadOPRR_sp
                                    ft_OPRR.Post();
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("OPRR", ex.Message);
                                }
                                break;
                            }
                        case "Update Pick List":
                            {
                                try
                                {
                                    ft_OPKL.Post();
                                }
                                catch (Exception ex)
                                {
                                    Log($"{ ex.Message } \n");
                                    ft_General.UpdateError("OPRR", ex.Message);
                                }
                                break;
                            }
                    }
                });
            }
            catch (Exception e)
            {
                ft_General.UpdateError("Main", e.Message);
                Log($"{e.Message}\n{ e.StackTrace}");
            }
        }
      
        void GenerateJsonAndSave(dynamic dataObject, Request request)
        {
            Log($"Request {request.request} by user {request.sapUser}, json infor = \n {JsonConvert.SerializeObject(dataObject)}");
        }

        void UpdateRequest(string guid, string status, string lastErrorMessag, string sapDocNumber)
        {
            try
            {
                var sql = @"UPDATE zmwRequest 
                            SET status = @status, 
                            lastErrorMessage = @lastErrorMessag, 
                            sapDocNumber = @sapDocNumber
                            WHERE GUID = @guid";

                var result = new SqlConnection(this.DbConnectString_Midware)
                    .Execute(sql,
                    new
                    {
                        status,
                        lastErrorMessag,
                        sapDocNumber,
                        guid
                    });
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{ e.StackTrace}");
            }
        }
    }
}
