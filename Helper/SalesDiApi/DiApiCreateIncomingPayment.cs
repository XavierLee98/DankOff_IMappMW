using Dapper;
using Microsoft.Data.SqlClient;
using IMAppSapMidware_NetCore.Models.IncomingPayment;
using IMAppSapMidware_NetCore.Models.Share;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IMAppSapMidware_NetCore.Helper.DiApi
{
    public class DiApiCreateIncomingPayment
    {
        public ErpPropertyHelper ErpProperty { get; set; } = null;
        public string LastSAPMsg { get; set; } = string.Empty;
        public string Midware_DBConnStr { get; set; } = string.Empty;
        public string Erp_DBConnStr { get; set; } = string.Empty;
        public string Guid { get; set; } = string.Empty;
        public string DocType { get; set; } = string.Empty;
        public string PostedDocNum { get; set; } = string.Empty;
        public int AttachedFileCnt { get; set; } = -1;
        public PaymentsDocHeader Payment { get; set; } = null;
        public List<PaymentsDocDetails> PaymentDetails { get; set; } = null;
        public List<PaymentsDocMeans> PaymentMeans { get; set; } = null;
        Company sapCompany { get; set; }

        void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }

        public bool CreateIncomingPaymentDocument()
        {
            string modName = $"[CreateIncomingPaymentDocument][{DocType}]";
            try
            {
                bool retResult = false;
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

                Payments py = (Payments)sapCompany.GetBusinessObject(BoObjectTypes.oPaymentsDrafts);
                py.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments;
                var curSapTableName = "OPDF"; // incoming payment draft table
                var curCardCode = FormatValue(Payment.CardCode, 15); //ut.SafeGetStr(ipHead.Rows[0]["cardCode"]), 15);
                py.CardCode = curCardCode;
                py.DocTypte = BoRcptTypes.rCustomer;
                py.DocDate = Payment.DocDate; // ut.SafeGetDateTime(ipHead.Rows[0]["docDate"]);
                py.DueDate = Payment.DueDate; //ut.SafeGetDateTime(ipHead.Rows[0]["dueDate"]);
                py.TaxDate = Payment.TaxDate; //ut.SafeGetDateTime(ipHead.Rows[0]["taxDate"]);
                py.HandWritten = BoYesNoEnum.tNO;
                py.JournalRemarks = FormatValue(Payment.JrnlMemo, 50); //ut.SafeGetStr(ipHead.Rows[0]["jrnlMemo"]), 50);
                py.Remarks = FormatValue(Payment.Comments, 254); //ut.SafeGetStr(ipHead.Rows[0]["comments"]), 254);
                py.Reference2 = FormatValue(Payment.Ref2, 11); //ut.SafeGetStr(ipHead.Rows[0]["ref2"]), 11);

                //? 20200315T1352
                // handle the doc currency (local or foreign)
                string cardCurrency = GetCardCurrency(curCardCode);
                string mainCurrency = GetCompanyMainCurreny();
                bool usedLocalCurrency = (cardCurrency.Equals(mainCurrency));

                py.DocCurrency = usedLocalCurrency ? mainCurrency : cardCurrency;
                //py.LocalCurrency = usedLocalCurrency ? BoYesNoEnum.tYES : BoYesNoEnum.tNO;

                // must be i local currency , will test on 16 Mar - Monday
                py.CashSum = RounToTwoDecimalPlace(Payment.CollectedCash); //ut.SafeGetDouble(ipHead.Rows[0]["CollectedCash"]);

                // User define field for integration 
                string docNumber = Payment.DocNum; // ut.SafeGetStr(ipHead.Rows[0]["docNum"]);
                py.UserFields.Fields.Item("U_Guid").Value = Guid; //guidVal;
                py.UserFields.Fields.Item("U_DocNumber").Value = docNumber;

                // for bank transfer mean update
                var transferSum = Payment.TransferSum; //ut.SafeGetDouble(ipHead.Rows[0]["TransferSum"]);
                var transferAcc = GetBankCodeGLAccount(Payment.TransferAccount); //ut.SafeGetStr(ipHead.Rows[0]["TransferAccount"]));
                var transferDate = Payment.TransDate;// ut.SafeGetDateTime(ipHead.Rows[0]["TransferDate"]);
                var transferReference = Payment.TransferReference; // ut.SafeGetStr(ipHead.Rows[0]["TransferReference"]);
                if (transferSum > 0 && transferAcc.Length > 0)
                {
                    py.TransferAccount = transferAcc;
                    py.TransferDate = transferDate;
                    py.TransferSum = RounToTwoDecimalPlace(transferSum);
                    py.TransferReference = transferReference;
                }

                /// add in the check payment 
                ///  ----------------------------------------------------

                int lineCounter = 0;
                if (PaymentDetails != null)
                {
                    PaymentDetails.ForEach(paydetail =>
                    {
                        py.Invoices.DocEntry = paydetail.DocEntry; // ut.SafeGetInt(inv["InvDocEntry"]);                        
                        py.Invoices.SetCurrentLine(lineCounter);

                        if (paydetail.DocType.Equals("IN"))
                        {
                            py.Invoices.InstallmentId = paydetail.InstlmntID; // ut.SafeGetInt(inv["InstllmntId"]);
                            py.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                            /// either applied the amount to foregn of local
                            double paymentAmt = paydetail.PaymentAmount;  //ut.SafeGetDouble(inv["PaymentAmount"]);
                            if (!usedLocalCurrency)
                            {
                                py.Invoices.AppliedFC = RounToTwoDecimalPlace(paydetail.PaymentAmount);
                            }
                            else
                            {
                                py.Invoices.SumApplied = RounToTwoDecimalPlace(paymentAmt);
                                py.Invoices.AppliedFC = 0;
                            }
                            py.Invoices.DiscountPercent = 0;
                        }
                        else
                        {
                            py.Invoices.InvoiceType = BoRcptInvTypes.it_CredItnote;
                            /// either applied the amount to foregn of local
                            double paymentAmt = paydetail.PaymentAmount;  //ut.SafeGetDouble(inv["PaymentAmount"]);
                            if (!usedLocalCurrency)
                            {
                                py.Invoices.AppliedFC = -1 * RounToTwoDecimalPlace(paydetail.PaymentAmount);
                            }
                            else
                            {
                                py.Invoices.SumApplied = -1 * RounToTwoDecimalPlace(paymentAmt);
                                py.Invoices.AppliedFC = 0;
                            }
                        }
                        py.Invoices.DocLine = lineCounter;
                        py.Invoices.Add();
                        lineCounter++;
                    });
                }


                lineCounter = 0;
                if (PaymentMeans != null)
                {
                    PaymentMeans.ForEach(mean =>
                    {
                        py.Checks.BankCode = mean.BankCode; //ut.SafeGetStr(mean["BankCode"]);

                        bool isNumeric = Int32.TryParse(mean.ChequeNum, out int result);
                        if (isNumeric)
                        {
                            py.Checks.CheckNumber = result;//  ut.SafeGetInt(mean["ChequeNum"]);
                        }

                        py.Checks.CheckSum = RounToTwoDecimalPlace(mean.Amount);// ut.SafeGetDouble(mean["Amount"]);
                        py.Checks.Details = "App Cheque Collected";
                        py.Checks.DueDate = mean.DueDate; // ut.SafeGetDateTime(mean["DueDate"]);
                        py.Checks.Add();
                        lineCounter++;
                    });
                }

                int addResult = py.Add();
                if (addResult == 0) // success
                {
                    var newKey = Convert.ToInt32(sapCompany.GetNewObjectKey());
                    var result = newKey.ToString(); // docentry from the object 
                    var doNum = GetDocNumberbyDoEntry(result, curSapTableName);
                    PostedDocNum = doNum.ToString();

                    // attche fle 
                    if (AttachedFileCnt > 0)
                    {
                        AddFileAttachment(newKey, BoObjectTypes.oPaymentsDrafts);
                    }

                    Log($"{modName} Created, DocEntry: {result} ,DocNo: {PostedDocNum}\n{Payment.Guid}");
                    return (!string.IsNullOrWhiteSpace(PostedDocNum));
                }
                else
                {
                    Log($"{sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
                }
                return retResult;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return false;
            }
        }

        double RounToTwoDecimalPlace(double mathVal)
        {
            return Math.Round(mathVal, 2);
        }

        void AddFileAttachment(int docEntry, BoObjectTypes objectType)
        {
            try
            {
                var files = GetAttachmentFiles(Guid);
                if (files == null) return;
                if (files.Count == 0) return;

                Documents doc = (Documents)sapCompany.GetBusinessObject(objectType);
                doc.GetByKey(docEntry);

                Attachments2 attachment = (Attachments2)sapCompany.GetBusinessObject(BoObjectTypes.oAttachments2);
                int fileCnt = 0;

                files.ForEach(file =>
                {
                    if (File.Exists(file.ServerSavedPath))
                    {
                        attachment.Lines.Add();
                        attachment.Lines.FileName = Path.GetFileNameWithoutExtension(file.ServerSavedPath);
                        attachment.Lines.FileExtension = Path.GetExtension(file.ServerSavedPath).Substring(1);
                        attachment.Lines.SourcePath = Path.GetDirectoryName(file.ServerSavedPath);
                        attachment.Lines.Override = BoYesNoEnum.tYES;
                        fileCnt++;
                    }
                });


                if (fileCnt != files.Count)
                {
                    Log($"Payment Doc Attachments {sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
                    return;
                }


                if (attachment.Add() == 0)
                {
                    int iAttEntry = int.Parse(sapCompany.GetNewObjectKey());
                    doc.AttachmentEntry = iAttEntry;
                    doc.Update();
                    return;
                }

                // else                                    
                Log($"Payment Doc Attachments {sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
            }
        }

        string FormatValue(string val, int requireLength)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(val))
                {
                    return string.Empty;
                }

                if (val.Length > requireLength)
                {
                    return val.Substring(0, requireLength);
                }
                return val;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return string.Empty;
            }
        }

        List<FileUpload> GetAttachmentFiles(string guid) // query from midware
        {
            try
            {
                var sqlQuery = @"SELECT * 
                                 FROM zmwFileUpload 
                                 WHERE HeaderGuid= @guid";

                return new SqlConnection(Midware_DBConnStr).Query<FileUpload>(sqlQuery, new { guid }).ToList();
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return null;
            }
        }


        string GetBankCodeGLAccount(string BankCode)
        {
            // SELECT GLAccount FROM ODSC t0 INNER JOIN DSC1 t1 ON t0.BankCode = t1.BankCode  WHERE t0.BankCode = 'PBB'

            try
            {
                string query = "SELECT GLAccount " +
                    "FROM ODSC t0 INNER JOIN DSC1 t1 ON t0.BankCode = t1.BankCode  " +
                    "WHERE t0.BankCode = @BankCode";

                return new SqlConnection(this.Erp_DBConnStr).ExecuteScalar<string>(query, new { BankCode });
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return string.Empty;
            }
        }

        int GetDocNumberbyDoEntry(string docEntry, string tableName)
        {
            try
            {
                string query = $"SELECT DocNum " +
                                $"FROM {tableName} " +
                                $"WHERE DocEntry=@docEntry";

                var result=  new SqlConnection(this.Erp_DBConnStr).ExecuteScalar<int>(query, new { docEntry });
                return result;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return -1;
            }
        }

        string GetCardCurrency(string cardCode)
        {
            try
            {
                string query = $"SELECT Currency " +
                                $"FROM OCRD " +
                                $"WHERE CardCode=@CardCode";

                var result = new SqlConnection(this.Erp_DBConnStr).ExecuteScalar<string>(query, new { @CardCode = cardCode });
                if (!string.IsNullOrWhiteSpace(result))
                {
                    return result;
                }
                return GetCompanyMainCurreny();
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return "MY";
            }
        }

        string GetCompanyMainCurreny()
        {
            try
            {
                var result = new SqlConnection(this.Erp_DBConnStr).ExecuteScalar<string>("SELECT TOP 1 MainCurncy FROM OADM");
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
