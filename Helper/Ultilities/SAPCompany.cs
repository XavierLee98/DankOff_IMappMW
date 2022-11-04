using Microsoft.Data.SqlClient;
using Newtonsoft.Json;
using IMAppSapMidware_NetCore.Models.SAPModels;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Reflection;

namespace IMAppSapMidware_NetCore
{
    public class SAPCompany
    {
        public Company oCom;
        public string errMsg { get; set; }
        public string UserID { get; set; }

        public SAPParam sapParam { get; set; }

        public SAPCompany() { }
        public bool connectSAP()
        {
            if (oCom != null)
            {
                if (oCom.Connected) return true;
                else
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCom);
                    oCom = null;
                }
            }
            oCom = new SAPbobsCOM.Company();
            oCom.DbServerType = (SAPbobsCOM.BoDataServerTypes)int.Parse(sapParam.SAPDBType);
            oCom.Server = sapParam.SAPServer;
            oCom.DbUserName = sapParam.DBUser;
            oCom.DbPassword = sapParam.DBPass;
            oCom.CompanyDB = sapParam.SAPCompany;
            oCom.UserName = sapParam.SAPUser;
            oCom.Password = sapParam.SAPPass;
            //oCom.LicenseServer = sapParam.SAPLicense;
            oCom.SLDServer = sapParam.SAPLicense;

            if (oCom.Connect() != 0)
            {
                errMsg = oCom.GetLastErrorDescription();
                oCom = null;
                return false;
            }

            return true;
        }
        private string getConnectionString()
        {
            string connstr = "server=" + sapParam.SAPServer + ";User ID=" + sapParam.DBUser + ";Password=" + sapParam.DBPass + ";database=" + sapParam.SAPCompany + ";MultipleActiveResultSets=True";//transaction binding=explicit Unbind ";
            return connstr;
        }
        private bool generateLog(object obj, string doctype, string logUser)
        {
            string jsonstr = JsonConvert.SerializeObject(obj);
            try
            {
                using (SqlConnection conn = new SqlConnection())
                {
                    conn.ConnectionString = ConfigurationManager.ConnectionStrings["AuthContext"].ConnectionString;
                    conn.Open();
                    using (SqlCommand comm = conn.CreateCommand())
                    {
                        comm.CommandText = "INSERT INTO [JSONREC]([AuthUser],[LogUserID],[ObjectType],[JSONObj],[UPDDATE])VALUES(@AuthUser,@LogUserID,@ObjectType,@JSONObj,GETDATE())";
                        comm.Parameters.Clear();
                        comm.Parameters.AddWithValue("@AuthUser", UserID);
                        comm.Parameters.AddWithValue("@LogUserID", logUser);
                        comm.Parameters.AddWithValue("@ObjectType", doctype);
                        comm.Parameters.AddWithValue("@JSONObj", jsonstr);
                        comm.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return false;
            }
            return true;
        }
        private bool generateLog(object obj, string doctype)
        {
            return generateLog(obj, doctype, "");
        }
        public int CreateAttachment(List<string> attachments)
        {
            int iAttEntry = -1;

            try
            {
                SAPbobsCOM.Attachments2 oATT = (SAPbobsCOM.Attachments2)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2);

                if (attachments.Count > 0)
                {
                    foreach (string attachment in attachments)
                    {
                        oATT.Lines.Add();
                        oATT.Lines.FileName = System.IO.Path.GetFileNameWithoutExtension(attachment);
                        oATT.Lines.FileExtension = System.IO.Path.GetExtension(attachment).Substring(1);
                        oATT.Lines.SourcePath = System.IO.Path.GetDirectoryName(attachment);
                        oATT.Lines.Override = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    int retcode = 0;
                    retcode = oATT.Add();
                    if (retcode == 0)
                    {
                        iAttEntry = int.Parse(oCom.GetNewObjectKey());
                        //Assign the attachment to the GR object (GR is my SAPbobsCOM.Documents object)
                    }
                    else
                    {
                        errMsg = oCom.GetLastErrorDescription();
                    }
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return iAttEntry;
            }
            return iAttEntry;

        }
        //public bool UpdateDocuments(Models.Documents doc, ref string docentry)
        //{
        //    oCom.StartTransaction();
        //    try
        //    {
        //        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), doc.DocObject));
        //        oDoc.GetByKey(int.Parse(docentry));
        //        if(!String.IsNullOrEmpty(doc.AttachmentEntry.ToString()) && doc.AttachmentEntry > 0) oDoc.AttachmentEntry = doc.AttachmentEntry;
        //        if (!String.IsNullOrEmpty(doc.Reference2)) oDoc.Reference2 = doc.Reference2;
        //        if (!String.IsNullOrEmpty(doc.JournalMemo)) oDoc.JournalMemo = doc.JournalMemo;
        //        if (!String.IsNullOrEmpty(doc.NumAtCard)) oDoc.NumAtCard = doc.NumAtCard;
        //        int retcode = oDoc.Update();
        //        if (retcode != 0)
        //        {
        //            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //            errMsg = oCom.GetLastErrorDescription();
        //            return false;
        //        }
        //        if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        errMsg = ex.Message;
        //        return false;
        //    }
        //}
        public bool CancelDocuments(Models.SAPModels.Documents doc, ref string docentry)
        {
            oCom.StartTransaction();

            try
            {
                SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), doc.DocObject));
                oDoc.GetByKey(int.Parse(docentry));
                oDoc.Comments = doc.Comments;
                int retcode = oDoc.Update();
                retcode = oDoc.Cancel();
                if (retcode != 0)
                {
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    errMsg = oCom.GetLastErrorDescription();
                    return false;
                }
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                return true;
            }
            catch (Exception ex)
            {
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                errMsg = ex.Message;
                return false;
            }
        }
        public bool CloseDocuments(Models.SAPModels.Documents doc, ref string docentry)
        {
            oCom.StartTransaction();

            try
            {
                SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), doc.DocObject));
                oDoc.GetByKey(int.Parse(docentry));
                oDoc.Comments = doc.Comments;
                int retcode = oDoc.Update();
                retcode = oDoc.Close();
                if (retcode != 0)
                {
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    errMsg = oCom.GetLastErrorDescription();
                    return false;
                }
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                return true;
            }
            catch (Exception ex)
            {
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                errMsg = ex.Message;
                return false;
            }
        }
        public bool CreateStockTransfer(Models.SAPModels.StockTransfer doc, ref string docentry)
        {
            oCom.StartTransaction();
            try
            {
                SAPbobsCOM.StockTransfer oDoc = (SAPbobsCOM.StockTransfer)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
                #region Header
                if (doc.Series > 0) oDoc.Series = doc.Series;
                oDoc.DocDate = doc.DocDate;
                if (!String.IsNullOrEmpty(doc.CardCode)) oDoc.CardCode = doc.CardCode;
                if (!String.IsNullOrEmpty(doc.Address)) oDoc.Address = doc.Address;
                if (!String.IsNullOrEmpty(doc.ATDocumentType)) oDoc.ATDocumentType = doc.ATDocumentType;
                if (!String.IsNullOrEmpty(doc.AuthorizationCode)) oDoc.AuthorizationCode = doc.AuthorizationCode;
                if (!String.IsNullOrEmpty(doc.Comments)) oDoc.Comments = doc.Comments;
                if (doc.FolioNumber > 0) oDoc.FolioNumber = doc.FolioNumber;
                if (!String.IsNullOrEmpty(doc.JournalMemo)) oDoc.JournalMemo = doc.JournalMemo;
                if (!String.IsNullOrEmpty(doc.Reference1)) oDoc.Reference1 = doc.Reference1;
                if (!String.IsNullOrEmpty(doc.Reference2)) oDoc.Reference2 = doc.Reference2;
                if (!String.IsNullOrEmpty(doc.FromWarehouse)) oDoc.FromWarehouse = doc.FromWarehouse;
                if (!String.IsNullOrEmpty(doc.ToWarehouse)) oDoc.ToWarehouse = doc.ToWarehouse;
                if (doc.TaxDate != DateTime.MinValue) oDoc.TaxDate = doc.TaxDate;

                foreach (PropertyInfo info in doc.UserFields.GetType().GetProperties())
                {
                    if (doc.UserFields.GetType().GetProperty(info.Name).GetValue(doc.UserFields) != null)
                    {
                        oDoc.UserFields.Fields.Item(info.Name).Value = doc.UserFields.GetType().GetProperty(info.Name).GetValue(doc.UserFields);
                    }
                }
                if (!String.IsNullOrEmpty(doc.U_UName)) oDoc.UserFields.Fields.Item("U_UName").Value = doc.U_UName;

                #endregion
                #region Details
                if (doc.Lines != null)
                {
                    for (int i = 0; i < doc.Lines.Count; i++)
                    {
                        if (i > 0) oDoc.Lines.Add();
                        oDoc.Lines.SetCurrentLine(i);
                        if (!String.IsNullOrEmpty(doc.Lines[i].ItemCode)) oDoc.Lines.ItemCode = doc.Lines[i].ItemCode;
                        if (!String.IsNullOrEmpty(doc.Lines[i].ItemDescription)) oDoc.Lines.ItemDescription = doc.Lines[i].ItemDescription;
                        if (doc.Lines[i].Price > 0) oDoc.Lines.Price = doc.Lines[i].Price;
                        if (!String.IsNullOrEmpty(doc.Lines[i].ProjectCode)) oDoc.Lines.ProjectCode = doc.Lines[i].ProjectCode;
                        if (doc.Lines[i].Quantity > 0) oDoc.Lines.Quantity = doc.Lines[i].Quantity;
                        if (doc.Lines[i].UnitPrice > 0) oDoc.Lines.UnitPrice = doc.Lines[i].UnitPrice;
                        if (doc.Lines[i].BaseEntry > 0)
                        {
                            oDoc.Lines.BaseEntry = doc.Lines[i].BaseEntry;
                            oDoc.Lines.BaseLine = doc.Lines[i].BaseLine;
                        }
                        if (!String.IsNullOrEmpty(doc.Lines[i].FromWarehouseCode)) oDoc.Lines.FromWarehouseCode = doc.Lines[i].FromWarehouseCode;
                        if (!String.IsNullOrEmpty(doc.Lines[i].WarehouseCode)) oDoc.Lines.WarehouseCode = doc.Lines[i].WarehouseCode;
                        if (!String.IsNullOrEmpty(doc.Lines[i].DistributionRule)) oDoc.Lines.DistributionRule = doc.Lines[i].DistributionRule;
                        if (!String.IsNullOrEmpty(doc.Lines[i].DistributionRule2)) oDoc.Lines.DistributionRule2 = doc.Lines[i].DistributionRule2;
                        if (!String.IsNullOrEmpty(doc.Lines[i].DistributionRule3)) oDoc.Lines.DistributionRule3 = doc.Lines[i].DistributionRule3;
                        if (!String.IsNullOrEmpty(doc.Lines[i].DistributionRule4)) oDoc.Lines.DistributionRule4 = doc.Lines[i].DistributionRule4;
                        if (!String.IsNullOrEmpty(doc.Lines[i].DistributionRule5)) oDoc.Lines.DistributionRule5 = doc.Lines[i].DistributionRule5;
                        if (doc.Lines[i].DiscountPercent != 0) oDoc.Lines.DiscountPercent = doc.Lines[i].DiscountPercent;


                        foreach (PropertyInfo info in doc.Lines[i].UserFields.GetType().GetProperties())
                        {
                            if (doc.Lines[i].UserFields.GetType().GetProperty(info.Name).GetValue(doc.Lines[i].UserFields) != null)
                            {
                                oDoc.Lines.UserFields.Fields.Item(info.Name).Value = doc.Lines[i].UserFields.GetType().GetProperty(info.Name).GetValue(doc.Lines[i].UserFields);
                            }
                        }
                    }
                }
                #endregion


                if (oDoc.Add() != 0)
                {
                    errMsg = oCom.GetLastErrorDescription();
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oDoc = null;
                    return false;
                }
                oCom.GetNewObjectCode(out docentry);
                if (docentry == "")
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oDoc = null;
                    return false;
                }

                SAPbobsCOM.StockTransfer getDoc = (SAPbobsCOM.StockTransfer)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

                if (!getDoc.GetByKey(int.Parse(docentry)))
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oDoc = null;
                    return false;
                }
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(getDoc);
                oDoc = null;
                getDoc = null;
            }
            catch (Exception ex)
            {
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                errMsg = ex.Message;
                return false;
            }
            return true;
        }
        public bool CreateDocuments(Models.SAPModels.Documents doc, ref string docentry)
        {
            //if (!generateLog(doc, doc.DocObject, doc.LogUserID)) return false;

            oCom.StartTransaction();
            try
            {
                SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), doc.DocObject));
                SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                #region Header
                if (doc.HandWritten)
                {
                    oDoc.DocNum = doc.DocNum;
                    oDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES;
                }
                else
                {
                    if (doc.Series > 0) oDoc.Series = doc.Series;
                }
                oDoc.DocDate = doc.DocDate;

                if (!String.IsNullOrEmpty(doc.CardCode)) oDoc.CardCode = doc.CardCode;
                if (!String.IsNullOrEmpty(doc.Address)) oDoc.Address = doc.Address;
                if (!String.IsNullOrEmpty(doc.Address2)) oDoc.Address2 = doc.Address2;
                if (!String.IsNullOrEmpty(doc.AgentCode)) oDoc.AgentCode = doc.AgentCode;
                if (doc.ApplyCurrentVATRatesForDownPaymentsToDraw) oDoc.ApplyCurrentVATRatesForDownPaymentsToDraw = SAPbobsCOM.BoYesNoEnum.tYES; else oDoc.ApplyCurrentVATRatesForDownPaymentsToDraw = SAPbobsCOM.BoYesNoEnum.tNO;
                if (doc.ApplyTaxOnFirstInstallment) oDoc.ApplyTaxOnFirstInstallment = SAPbobsCOM.BoYesNoEnum.tYES; else oDoc.ApplyTaxOnFirstInstallment = SAPbobsCOM.BoYesNoEnum.tNO;
                if (doc.ArchiveNonremovableSalesQuotation) oDoc.ArchiveNonremovableSalesQuotation = SAPbobsCOM.BoYesNoEnum.tYES; else oDoc.ArchiveNonremovableSalesQuotation = SAPbobsCOM.BoYesNoEnum.tNO;
                if (doc.AssetValueDate != DateTime.MinValue) oDoc.AssetValueDate = doc.AssetValueDate;
                if (!String.IsNullOrEmpty(doc.ATDocumentType)) oDoc.ATDocumentType = doc.ATDocumentType;
                if (!String.IsNullOrEmpty(doc.AuthorizationCode)) oDoc.AuthorizationCode = doc.AuthorizationCode;
                if (doc.BlanketAgreementNumber > 0) oDoc.BlanketAgreementNumber = doc.BlanketAgreementNumber;
                if (doc.BlockDunning) oDoc.BlockDunning = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(doc.Box1099)) oDoc.Box1099 = doc.Box1099;
                if (!String.IsNullOrEmpty(doc.BPChannelCode)) oDoc.BPChannelCode = doc.BPChannelCode;
                if (doc.BPChannelContact > 0) oDoc.BPChannelContact = doc.BPChannelContact;
                if (doc.BPL_IDAssignedToInvoice > 0) oDoc.BPL_IDAssignedToInvoice = doc.BPL_IDAssignedToInvoice;
                if (doc.CashDiscountDateOffset > 0) oDoc.CashDiscountDateOffset = doc.CashDiscountDateOffset;
                if (!String.IsNullOrEmpty(doc.CentralBankIndicator)) oDoc.CentralBankIndicator = doc.CentralBankIndicator;
                if (doc.ClosingDate != DateTime.MinValue) oDoc.ClosingDate = doc.ClosingDate;
                if (!String.IsNullOrEmpty(doc.ClosingRemarks)) oDoc.ClosingRemarks = doc.ClosingRemarks;
                if (!String.IsNullOrEmpty(doc.Comments)) oDoc.Comments = doc.Comments;
                if (doc.ContactPersonCode > 0) oDoc.ContactPersonCode = doc.ContactPersonCode;
                if (doc.DeferredTax) oDoc.DeferredTax = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.DiscountPercent != 0) oDoc.DiscountPercent = doc.DiscountPercent;
                if (!String.IsNullOrEmpty(doc.DocCurrency)) oDoc.DocCurrency = doc.DocCurrency;
                if (doc.DocDueDate != DateTime.MinValue) oDoc.DocDueDate = doc.DocDueDate;
                if (doc.DocRate > 0) oDoc.DocRate = doc.DocRate;
                if (doc.DocumentsOwner > 0) oDoc.DocumentsOwner = doc.DocumentsOwner;
                if (doc.DocType > 0) oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service; else oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                if (doc.FolioNumber > 0) oDoc.FolioNumber = doc.FolioNumber;
                if (doc.ImportFileNum > 0) oDoc.ImportFileNum = doc.ImportFileNum;
                if (doc.InsuranceOperation347) oDoc.InsuranceOperation347 = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(doc.JournalMemo)) oDoc.JournalMemo = doc.JournalMemo;
                if (!String.IsNullOrEmpty(doc.ManualNumber)) oDoc.ManualNumber = doc.ManualNumber;
                if (doc.LanguageCode > 0) oDoc.LanguageCode = doc.LanguageCode;
                if (doc.MaximumCashDiscount) oDoc.MaximumCashDiscount = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.NTSApproved)
                {
                    oDoc.NTSApproved = SAPbobsCOM.BoYesNoEnum.tYES;
                    oDoc.NTSApprovedNumber = doc.NTSApprovedNumber;
                }
                if (!String.IsNullOrEmpty(doc.NumAtCard)) oDoc.NumAtCard = doc.NumAtCard;
                if (doc.NumberOfInstallments > 0) oDoc.NumberOfInstallments = doc.NumberOfInstallments;
                if (doc.OpenForLandedCosts) oDoc.OpenForLandedCosts = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(doc.OpeningRemarks)) oDoc.OpeningRemarks = doc.OpeningRemarks;
                if (doc.PartialSupply) oDoc.PartialSupply = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.PaymentBlock) oDoc.PaymentBlock = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.PaymentBlockEntry > 0) oDoc.PaymentBlockEntry = doc.PaymentBlockEntry;
                if (doc.PaymentGroupCode > 0) oDoc.PaymentGroupCode = doc.PaymentGroupCode;
                if (!String.IsNullOrEmpty(doc.PaymentMethod)) oDoc.PaymentMethod = doc.PaymentMethod;
                if (!String.IsNullOrEmpty(doc.PaymentReference)) oDoc.PaymentReference = doc.PaymentReference;
                if (!String.IsNullOrEmpty(doc.PayToBankAccountNo)) oDoc.PayToBankAccountNo = doc.PayToBankAccountNo;
                if (!String.IsNullOrEmpty(doc.PayToBankBranch)) oDoc.PayToBankBranch = doc.PayToBankBranch;
                if (!String.IsNullOrEmpty(doc.PayToBankCode)) oDoc.PayToBankCode = doc.PayToBankCode;
                if (!String.IsNullOrEmpty(doc.PayToBankCountry)) oDoc.PayToBankCountry = doc.PayToBankCountry;
                if (!String.IsNullOrEmpty(doc.PayToCode)) oDoc.PayToCode = doc.PayToCode;
                if (doc.POSCashierNumber > 0) oDoc.POSCashierNumber = doc.POSCashierNumber;
                if (!String.IsNullOrEmpty(doc.POSEquipmentNumber)) oDoc.POSEquipmentNumber = doc.POSEquipmentNumber;
                if (!String.IsNullOrEmpty(doc.POSManufacturerSerialNumber)) oDoc.POSManufacturerSerialNumber = doc.POSManufacturerSerialNumber;
                if (!String.IsNullOrEmpty(doc.Project)) oDoc.Project = doc.Project;
                if (doc.Receiver > 0) oDoc.Receiver = doc.Receiver;
                if (!String.IsNullOrEmpty(doc.Reference1)) oDoc.Reference1 = doc.Reference1;
                if (!String.IsNullOrEmpty(doc.Reference2)) oDoc.Reference2 = doc.Reference2;
                if (doc.Releaser > 0) oDoc.Releaser = doc.Releaser;
                if (doc.RelevantToGTS) oDoc.RelevantToGTS = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.ReopenManuallyClosedOrCanceledDocument) oDoc.ReopenManuallyClosedOrCanceledDocument = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.ReopenOriginalDocument) oDoc.ReopenOriginalDocument = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.ReqType > 0) oDoc.ReqType = doc.ReqType;
                if (!String.IsNullOrEmpty(doc.Requester)) oDoc.Requester = doc.Requester;
                if (doc.RequesterBranch > 0) oDoc.RequesterBranch = doc.RequesterBranch;
                if (doc.RequesterDepartment > 0) oDoc.RequesterDepartment = doc.RequesterDepartment;
                if (!String.IsNullOrEmpty(doc.RequesterEmail)) oDoc.RequesterEmail = doc.RequesterEmail;
                if (!String.IsNullOrEmpty(doc.RequesterName)) oDoc.RequesterName = doc.RequesterName;
                if (doc.RequriedDate != DateTime.MinValue) oDoc.RequriedDate = doc.RequriedDate;
                if (doc.ReserveInvoice) oDoc.ReserveInvoice = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.RevisionPo) oDoc.RevisionPo = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.Rounding) oDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.RoundingDiffAmount > 0) oDoc.RoundingDiffAmount = doc.RoundingDiffAmount;
                if (doc.SalesPersonCode > 0) oDoc.SalesPersonCode = doc.SalesPersonCode;
                if (doc.SendNotification) oDoc.SendNotification = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.SequenceCode > 0) oDoc.SequenceCode = doc.SequenceCode;
                if (!String.IsNullOrEmpty(doc.SequenceModel)) oDoc.SequenceModel = doc.SequenceModel;
                if (doc.SequenceSerial > 0) oDoc.SequenceSerial = doc.SequenceSerial;
                if (!String.IsNullOrEmpty(doc.SeriesString)) oDoc.SeriesString = doc.SeriesString;
                if (doc.ServiceGrossProfitPercent > 0) oDoc.ServiceGrossProfitPercent = doc.ServiceGrossProfitPercent;
                if (!string.IsNullOrEmpty(doc.ShipToCode)) oDoc.ShipToCode = doc.ShipToCode;
                if (doc.ShowSCN) oDoc.ShowSCN = SAPbobsCOM.BoYesNoEnum.tYES;
                if (doc.SpecifiedClosingDate != DateTime.MinValue) oDoc.SpecifiedClosingDate = doc.SpecifiedClosingDate;
                if (!String.IsNullOrEmpty(doc.SubSeriesString)) oDoc.SubSeriesString = doc.SubSeriesString;
                if (!String.IsNullOrEmpty(doc.Supplier)) oDoc.Supplier = doc.Supplier;
                if (doc.TaxDate != DateTime.MinValue) oDoc.TaxDate = doc.TaxDate;
                if (!String.IsNullOrEmpty(doc.TaxExemptionLetterNum)) oDoc.TaxExemptionLetterNum = doc.TaxExemptionLetterNum;
                if (!String.IsNullOrEmpty(doc.TrackingNumber)) oDoc.TrackingNumber = doc.TrackingNumber;
                if (doc.TransportationCode > 0) oDoc.TransportationCode = doc.TransportationCode;
                if (doc.UseCorrectionVATGroup) oDoc.UseCorrectionVATGroup = SAPbobsCOM.BoYesNoEnum.tYES;

                if (!String.IsNullOrEmpty(doc.U_PRID)) oDoc.UserFields.Fields.Item("U_PRID").Value = doc.U_PRID;
                if (!String.IsNullOrEmpty(doc.U_POID)) oDoc.UserFields.Fields.Item("U_POID").Value = doc.U_POID;
                if (!String.IsNullOrEmpty(doc.U_GRNID)) oDoc.UserFields.Fields.Item("U_GRNID").Value = doc.U_GRNID;
                if (!String.IsNullOrEmpty(doc.U_GRTNID)) oDoc.UserFields.Fields.Item("U_GRTNID").Value = doc.U_GRTNID;
                if (!String.IsNullOrEmpty(doc.U_APINVID)) oDoc.UserFields.Fields.Item("U_APINVID").Value = doc.U_APINVID;
                if (!String.IsNullOrEmpty(doc.U_RefNo)) oDoc.UserFields.Fields.Item("U_RefNo").Value = doc.U_RefNo;
                if (!String.IsNullOrEmpty(doc.U_Requester)) oDoc.UserFields.Fields.Item("U_Requester").Value = doc.U_Requester;
                if (!String.IsNullOrEmpty(doc.U_OU)) oDoc.UserFields.Fields.Item("U_OU").Value = doc.U_OU;

                foreach (PropertyInfo info in doc.UserFields.GetType().GetProperties())
                {
                    if (doc.UserFields.GetType().GetProperty(info.Name).GetValue(doc.UserFields) != null)
                    {
                        oDoc.UserFields.Fields.Item(info.Name).Value = doc.UserFields.GetType().GetProperty(info.Name).GetValue(doc.UserFields);
                    }
                }

                if (!String.IsNullOrEmpty(doc.AttachmentEntry.ToString()) && doc.AttachmentEntry > 0) oDoc.AttachmentEntry = doc.AttachmentEntry;


                #endregion
                #region Details
                if (doc.Lines != null)
                {
                    for (int i = 0; i < doc.Lines.Count; i++)
                    {

                        if (i > 0) oDoc.Lines.Add();
                        oDoc.Lines.SetCurrentLine(i);
                        if (!String.IsNullOrEmpty(doc.Lines[i].AccountCode)) oDoc.Lines.AccountCode = doc.Lines[i].AccountCode;
                        if (doc.Lines[i].AgreementNo > 0)
                        {
                            oDoc.Lines.AgreementNo = doc.Lines[i].AgreementNo;
                            oDoc.Lines.AgreementRowNumber = doc.Lines[i].AgreementRowNumber;
                        }

                        if (doc.Lines[i].BackOrder) oDoc.Lines.BackOrder = SAPbobsCOM.BoYesNoEnum.tYES;
                        if (doc.Lines[i].BaseEntry > 0)
                        {
                            oDoc.Lines.BaseEntry = doc.Lines[i].BaseEntry;
                            oDoc.Lines.BaseLine = doc.Lines[i].BaseLine;
                            oDoc.Lines.BaseType = doc.Lines[i].BaseType;
                        }

                        if (!String.IsNullOrEmpty(doc.Lines[i].LineVendor)) oDoc.Lines.LineVendor = doc.Lines[i].LineVendor;
                        if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode;
                        if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode2)) oDoc.Lines.CostingCode2 = doc.Lines[i].CostingCode2;
                        if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode3)) oDoc.Lines.CostingCode3 = doc.Lines[i].CostingCode3;
                        if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode4)) oDoc.Lines.CostingCode4 = doc.Lines[i].CostingCode4;
                        if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode5)) oDoc.Lines.CostingCode5 = doc.Lines[i].CostingCode5;
                        if (doc.Lines[i].DiscountPercent != 0) oDoc.Lines.DiscountPercent = doc.Lines[i].DiscountPercent;
                        if (!String.IsNullOrEmpty(doc.Lines[i].ItemCode)) oDoc.Lines.ItemCode = doc.Lines[i].ItemCode;
                        if (!String.IsNullOrEmpty(doc.Lines[i].ItemDescription)) oDoc.Lines.ItemDescription = doc.Lines[i].ItemDescription;
                        if (!String.IsNullOrEmpty(doc.Lines[i].ItemDetails)) oDoc.Lines.ItemDetails = doc.Lines[i].ItemDetails;
                        //if (doc.Lines[i].LineTotal > 0 && doc.DocType > 0)
                        oDoc.Lines.LineTotal = doc.Lines[i].LineTotal;
                        if (doc.Lines[i].LocationCode > 0) oDoc.Lines.LocationCode = doc.Lines[i].LocationCode;
                        //if (doc.Lines[i].Price > 0)
                        //    if (!String.IsNullOrEmpty(doc.Lines[i].Price.ToString())) oDoc.Lines.Price = doc.Lines[i].Price;
                        //if (doc.Lines[i].UnitPrice > 0)
                        if (!String.IsNullOrEmpty(doc.Lines[i].UnitPrice.ToString())) oDoc.Lines.UnitPrice = doc.Lines[i].UnitPrice;
                        if (!String.IsNullOrEmpty(doc.Lines[i].WarehouseCode)) oDoc.Lines.WarehouseCode = doc.Lines[i].WarehouseCode;
                        if (doc.Lines[i].PriceAfterVAT > 0) oDoc.Lines.PriceAfterVAT = doc.Lines[i].PriceAfterVAT;
                        if (doc.Lines[i].GrossTotal > 0) oDoc.Lines.GrossTotal = doc.Lines[i].GrossTotal;
                        if (!String.IsNullOrEmpty(doc.Lines[i].ProjectCode)) oDoc.Lines.ProjectCode = doc.Lines[i].ProjectCode;
                        if (!String.IsNullOrEmpty(doc.Lines[i].Quantity.ToString())) oDoc.Lines.Quantity = doc.Lines[i].Quantity;
                        if (doc.Lines[i].RequiredDate != DateTime.MinValue) oDoc.Lines.RequiredDate = doc.Lines[i].RequiredDate;
                        if (doc.Lines[i].ShipDate != DateTime.MinValue) oDoc.Lines.ShipDate = doc.Lines[i].ShipDate;
                        if (doc.Lines[i].ShippingMethod > 0) oDoc.Lines.ShippingMethod = doc.Lines[i].ShippingMethod;
                        if (!String.IsNullOrEmpty(doc.Lines[i].SWW)) oDoc.Lines.SWW = doc.Lines[i].SWW;
                        if (!String.IsNullOrEmpty(doc.Lines[i].TaxCode)) oDoc.Lines.TaxCode = doc.Lines[i].TaxCode;
                        if (doc.Lines[i].TaxOnly) oDoc.Lines.TaxOnly = SAPbobsCOM.BoYesNoEnum.tYES;
                        if (doc.Lines[i].TaxTotal > 0) oDoc.Lines.TaxTotal = doc.Lines[i].TaxTotal;
                        if (doc.Lines[i].UnitPrice > 0) oDoc.Lines.UnitPrice = doc.Lines[i].UnitPrice;
                        if (!String.IsNullOrEmpty(doc.Lines[i].VatGroup)) oDoc.Lines.VatGroup = doc.Lines[i].VatGroup;
                        if (!String.IsNullOrEmpty(doc.Lines[i].U_PCode)) oDoc.Lines.UserFields.Fields.Item("U_PCode").Value = doc.Lines[i].U_PCode;
                        if (!String.IsNullOrEmpty(doc.Lines[i].U_ItemDesc)) oDoc.Lines.UserFields.Fields.Item("U_ItemDesc").Value = doc.Lines[i].U_ItemDesc;
                        if (!String.IsNullOrEmpty(doc.Lines[i].U_BarCode)) oDoc.Lines.UserFields.Fields.Item("U_BarCode").Value = doc.Lines[i].U_BarCode;


                        foreach (PropertyInfo info in doc.Lines[i].UserFields.GetType().GetProperties())
                        {
                            if (doc.Lines[i].UserFields.GetType().GetProperty(info.Name).GetValue(doc.Lines[i].UserFields) != null)
                            {
                                oDoc.Lines.UserFields.Fields.Item(info.Name).Value = doc.Lines[i].UserFields.GetType().GetProperty(info.Name).GetValue(doc.Lines[i].UserFields);
                            }
                        }

                        #region Serial Numbers
                        if (doc.Lines[i].Serials != null)
                        {
                            for (int j = 0; j < doc.Lines[i].Serials.Count; j++)
                            {
                                if (j > 0) oDoc.Lines.SerialNumbers.Add();
                                oDoc.Lines.SerialNumbers.SetCurrentLine(j);
                                if (!String.IsNullOrEmpty(doc.Lines[i].Serials[j].InternalSerialNumber)) oDoc.Lines.SerialNumbers.InternalSerialNumber = doc.Lines[i].Serials[j].InternalSerialNumber;
                                if (!String.IsNullOrEmpty(doc.Lines[i].Serials[j].ManufacturerSerialNumber)) oDoc.Lines.SerialNumbers.ManufacturerSerialNumber = doc.Lines[i].Serials[j].ManufacturerSerialNumber;
                                if (doc.Lines[i].Serials[j].ExpiryDate != DateTime.MinValue) oDoc.Lines.SerialNumbers.ExpiryDate = doc.Lines[i].Serials[j].ExpiryDate;
                                if (!String.IsNullOrEmpty(doc.Lines[i].Serials[j].Location)) oDoc.Lines.SerialNumbers.Location = doc.Lines[i].Serials[j].Location;
                                if (doc.Lines[i].Serials[j].ManufactureDate != DateTime.MinValue) oDoc.Lines.SerialNumbers.ManufactureDate = doc.Lines[i].Serials[j].ManufactureDate;
                                if (!String.IsNullOrEmpty(doc.Lines[i].Serials[j].Notes)) oDoc.Lines.SerialNumbers.Notes = doc.Lines[i].Serials[j].Notes;
                                if (doc.Lines[i].Serials[j].WarrantyStart != DateTime.MinValue) oDoc.Lines.SerialNumbers.WarrantyStart = doc.Lines[i].Serials[j].WarrantyStart;
                                if (doc.Lines[i].Serials[j].WarrantyEnd != DateTime.MinValue) oDoc.Lines.SerialNumbers.WarrantyEnd = doc.Lines[i].Serials[j].WarrantyEnd;

                                foreach (PropertyInfo info in doc.Lines[i].Serials[j].UserFields.GetType().GetProperties())
                                {
                                    if (doc.Lines[i].Serials[j].UserFields.GetType().GetProperty(info.Name).GetValue(doc.Lines[i].Serials[j].UserFields) != null)
                                    {
                                        oDoc.Lines.SerialNumbers.UserFields.Fields.Item(info.Name).Value = doc.Lines[i].Serials[j].UserFields.GetType().GetProperty(info.Name).GetValue(doc.Lines[i].Serials[j].UserFields);
                                    }
                                }
                            }
                        }
                        #endregion

                        #region Batch Numbers
                        if (doc.Lines[i].Batches != null)
                        {
                            for (int j = 0; j < doc.Lines[i].Batches.Count; j++)
                            {
                                if (j > 0) oDoc.Lines.BatchNumbers.Add();
                                oDoc.Lines.BatchNumbers.SetCurrentLine(j);
                                if (!String.IsNullOrEmpty(doc.Lines[i].Batches[j].InternalSerialNumber)) oDoc.Lines.BatchNumbers.InternalSerialNumber = doc.Lines[i].Batches[j].InternalSerialNumber;
                                if (!String.IsNullOrEmpty(doc.Lines[i].Batches[j].BatchNumber)) oDoc.Lines.BatchNumbers.BatchNumber = doc.Lines[i].Batches[j].BatchNumber;
                                if (doc.Lines[i].Batches[j].Quantity > 0) oDoc.Lines.BatchNumbers.Quantity = doc.Lines[i].Batches[j].Quantity;
                                if (doc.Lines[i].Batches[j].AddmisionDate != DateTime.MinValue) oDoc.Lines.BatchNumbers.AddmisionDate = doc.Lines[i].Batches[j].AddmisionDate;
                                if (doc.Lines[i].Batches[j].ExpiryDate != DateTime.MinValue) oDoc.Lines.BatchNumbers.ExpiryDate = doc.Lines[i].Batches[j].ExpiryDate;
                                if (!String.IsNullOrEmpty(doc.Lines[i].Batches[j].Location)) oDoc.Lines.BatchNumbers.Location = doc.Lines[i].Batches[j].Location;
                                if (doc.Lines[i].Batches[j].ManufacturingDate != DateTime.MinValue) oDoc.Lines.BatchNumbers.ManufacturingDate = doc.Lines[i].Batches[j].ManufacturingDate;
                                if (!String.IsNullOrEmpty(doc.Lines[i].Batches[j].ManufacturerSerialNumber)) oDoc.Lines.BatchNumbers.ManufacturerSerialNumber = doc.Lines[i].Batches[j].ManufacturerSerialNumber;
                                if (!String.IsNullOrEmpty(doc.Lines[i].Batches[j].Notes)) oDoc.Lines.BatchNumbers.Notes = doc.Lines[i].Batches[j].Notes;

                                foreach (PropertyInfo info in doc.Lines[i].Batches[j].UserFields.GetType().GetProperties())
                                {
                                    if (doc.Lines[i].Batches[j].UserFields.GetType().GetProperty(info.Name).GetValue(doc.Lines[i].Batches[j].UserFields) != null)
                                    {
                                        oDoc.Lines.BatchNumbers.UserFields.Fields.Item(info.Name).Value = doc.Lines[i].Batches[j].UserFields.GetType().GetProperty(info.Name).GetValue(doc.Lines[i].Batches[j].UserFields);
                                    }
                                }
                            }
                        }

                        #endregion
                    }
                }
                #endregion
                #region Downpayment to draw
                if (doc.DownPaymentToDraw != null)
                {
                    for (int i = 0; i < doc.DownPaymentToDraw.Count; i++)
                    {
                        if (i > 0) oDoc.DownPaymentsToDraw.Add();
                        oDoc.DownPaymentsToDraw.SetCurrentLine(i);

                        if (doc.DownPaymentToDraw[i].AmountToDraw > 0) oDoc.DownPaymentsToDraw.AmountToDraw = doc.DownPaymentToDraw[i].AmountToDraw;
                        if (doc.DownPaymentToDraw[i].AmountToDrawFC > 0) oDoc.DownPaymentsToDraw.AmountToDrawFC = doc.DownPaymentToDraw[i].AmountToDrawFC;
                        if (doc.DownPaymentToDraw[i].DocEntry > 0) oDoc.DownPaymentsToDraw.DocEntry = doc.DownPaymentToDraw[i].DocEntry;
                        if (doc.DownPaymentToDraw[i].GrossAmountToDraw > 0) oDoc.DownPaymentsToDraw.GrossAmountToDraw = doc.DownPaymentToDraw[i].GrossAmountToDraw;
                        if (doc.DownPaymentToDraw[i].GrossAmountToDrawFC > 0) oDoc.DownPaymentsToDraw.GrossAmountToDrawFC = doc.DownPaymentToDraw[i].GrossAmountToDrawFC;

                    }
                }
                #endregion

                if (oDoc.Add() != 0)
                {
                    errMsg = oCom.GetLastErrorDescription();
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oDoc = null;
                    return false;
                }
                oCom.GetNewObjectCode(out docentry);
                if (docentry == "")
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oDoc = null;
                    return false;
                }

                SAPbobsCOM.Documents getDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), doc.DocObject));

                if (!getDoc.GetByKey(int.Parse(docentry)))
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oDoc = null;
                    return false;
                }
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(getDoc);
                oDoc = null;
                getDoc = null;
            }
            catch (Exception ex)
            {
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                errMsg = ex.Message;
                return false;
            }
            return true;
        }

        public SAPbobsCOM.Documents getDocuments(string docEntry, string ObjectType)
        {
            if (docEntry == "") return null;
            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), ObjectType));
            int entryNo = 0;
            if (!int.TryParse(docEntry, out entryNo)) return null;
            if (oDoc.GetByKey(entryNo))
            {
                return oDoc;
            }

            return null;
        }
        public bool CreatePayments(Models.SAPModels.Payments pay, ref string docentry)
        {
            //if (!generateLog(pay, pay.DocObject, pay.LogUserID)) return false;
            oCom.StartTransaction();
            try
            {
                SAPbobsCOM.Payments oPay = (SAPbobsCOM.Payments)oCom.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), pay.DocObject));
                #region Header
                if (pay.HandWritten)
                {
                    oPay.DocNum = pay.DocNum;
                    oPay.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES;
                }
                else
                {
                    if (pay.Series > 0) oPay.Series = pay.Series;
                }
                if (!String.IsNullOrEmpty(pay.Address)) oPay.Address = pay.Address;
                if (pay.ApplyVAT) oPay.ApplyVAT = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(pay.BankAccount)) oPay.BankAccount = pay.BankAccount;
                if (pay.BankChargeAmount != 0) oPay.BankChargeAmount = pay.BankChargeAmount;
                if (!String.IsNullOrEmpty(pay.BankCode)) oPay.BankCode = pay.BankCode;
                if (pay.BPLID != 0) oPay.BPLID = pay.BPLID;
                if (!String.IsNullOrEmpty(pay.CardCode)) oPay.CardCode = pay.CardCode;
                if (!String.IsNullOrEmpty(pay.CashAccount)) oPay.CashAccount = pay.CashAccount;
                if (pay.CashSum != 0) oPay.CashSum = pay.CashSum;
                if (!String.IsNullOrEmpty(pay.CheckAccount)) oPay.CheckAccount = pay.CheckAccount;
                if (pay.ContactPersonCode != 0) oPay.ContactPersonCode = pay.ContactPersonCode;
                if (!String.IsNullOrEmpty(pay.ControlAccount)) oPay.ControlAccount = pay.ControlAccount;
                if (!String.IsNullOrEmpty(pay.CounterReference)) oPay.CounterReference = pay.CounterReference;
                if (pay.DeductionPercent != 0) oPay.DeductionPercent = pay.DeductionPercent;
                if (pay.DeductionSum != 0) oPay.DeductionSum = pay.DeductionSum;
                if (!String.IsNullOrEmpty(pay.DocCurrency)) oPay.DocCurrency = pay.DocCurrency;
                if (pay.DocDate != DateTime.MinValue) oPay.DocDate = pay.DocDate;
                if (pay.DocObject == "24") oPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments; else oPay.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_OutgoingPayments;
                oPay.DocType = (SAPbobsCOM.BoRcptTypes)pay.DocType;
                if (pay.DueDate != DateTime.MinValue) oPay.DueDate = pay.DueDate;
                if (pay.IsPayToBank) oPay.IsPayToBank = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(pay.JournalRemarks)) oPay.JournalRemarks = pay.JournalRemarks;
                if (!pay.LocalCurrency) oPay.LocalCurrency = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(pay.PayToBankAccountNo)) oPay.PayToBankAccountNo = pay.PayToBankAccountNo;
                if (!String.IsNullOrEmpty(pay.PayToBankBranch)) oPay.PayToBankBranch = pay.PayToBankBranch;
                if (!String.IsNullOrEmpty(pay.PayToBankCode)) oPay.PayToBankCode = pay.PayToBankCode;
                if (!String.IsNullOrEmpty(pay.PayToBankCountry)) oPay.PayToBankCountry = pay.PayToBankCountry;
                if (!String.IsNullOrEmpty(pay.PayToCode)) oPay.PayToCode = pay.PayToCode;
                if (!String.IsNullOrEmpty(pay.ProjectCode)) oPay.ProjectCode = pay.ProjectCode;
                if (!String.IsNullOrEmpty(pay.Reference1)) oPay.Reference1 = pay.Reference1;
                if (!String.IsNullOrEmpty(pay.Reference2)) oPay.Reference2 = pay.Reference2;
                if (!String.IsNullOrEmpty(pay.Remarks)) oPay.Remarks = pay.Remarks;
                if (pay.TaxDate != DateTime.MinValue) oPay.TaxDate = pay.TaxDate;
                if (!String.IsNullOrEmpty(pay.TaxGroup)) oPay.TaxGroup = pay.TaxGroup;
                if (!String.IsNullOrEmpty(pay.TransactionCode)) oPay.TransactionCode = pay.TransactionCode;
                if (!String.IsNullOrEmpty(pay.TransferAccount)) oPay.TransferAccount = pay.TransferAccount;
                if (pay.TransferDate != DateTime.MinValue) oPay.TransferDate = pay.TransferDate;
                if (!String.IsNullOrEmpty(pay.TransferReference)) oPay.TransferReference = pay.TransferReference;
                if (pay.TransferSum != 0) oPay.TransferSum = pay.TransferSum;
                if (pay.VatDate != DateTime.MinValue) oPay.VatDate = DateTime.Parse(pay.VatDate.ToString());

                foreach (PropertyInfo info in pay.UserFields.GetType().GetProperties())
                {
                    if (pay.UserFields.GetType().GetProperty(info.Name).GetValue(pay.UserFields) != null)
                    {
                        oPay.UserFields.Fields.Item(info.Name).Value = pay.UserFields.GetType().GetProperty(info.Name).GetValue(pay.UserFields);
                    }
                }
                #endregion
                #region Paid Invoices
                if (pay.AppliedInvoices != null)
                {
                    for (int i = 0; i < pay.AppliedInvoices.Count; i++)
                    {
                        if (i > 0) oPay.Invoices.Add();
                        oPay.Invoices.SetCurrentLine(i);
                        if (pay.AppliedInvoices[i].AppliedFC != 0) oPay.Invoices.AppliedFC = pay.AppliedInvoices[i].AppliedFC;
                        if (pay.AppliedInvoices[i].DiscountPercent != 0) oPay.Invoices.DiscountPercent = pay.AppliedInvoices[i].DiscountPercent;
                        if (!String.IsNullOrEmpty(pay.AppliedInvoices[i].DistributionRule)) oPay.Invoices.DistributionRule = pay.AppliedInvoices[i].DistributionRule;
                        if (!String.IsNullOrEmpty(pay.AppliedInvoices[i].DistributionRule2)) oPay.Invoices.DistributionRule2 = pay.AppliedInvoices[i].DistributionRule2;
                        if (!String.IsNullOrEmpty(pay.AppliedInvoices[i].DistributionRule3)) oPay.Invoices.DistributionRule3 = pay.AppliedInvoices[i].DistributionRule3;
                        if (!String.IsNullOrEmpty(pay.AppliedInvoices[i].DistributionRule4)) oPay.Invoices.DistributionRule4 = pay.AppliedInvoices[i].DistributionRule4;
                        if (!String.IsNullOrEmpty(pay.AppliedInvoices[i].DistributionRule5)) oPay.Invoices.DistributionRule5 = pay.AppliedInvoices[i].DistributionRule5;
                        oPay.Invoices.DocEntry = pay.AppliedInvoices[i].DocEntry;
                        if (pay.AppliedInvoices[i].DocLine != 0) oPay.Invoices.DocLine = pay.AppliedInvoices[i].DocLine;
                        oPay.Invoices.InvoiceType = (SAPbobsCOM.BoRcptInvTypes)pay.AppliedInvoices[i].InvoiceType;
                        if (pay.AppliedInvoices[i].SumApplied != 0) oPay.Invoices.SumApplied = pay.AppliedInvoices[i].SumApplied;
                        if (pay.AppliedInvoices[i].TotalDiscount != 0) oPay.Invoices.TotalDiscount = pay.AppliedInvoices[i].TotalDiscount;
                        if (pay.AppliedInvoices[i].TotalDiscountFC != 0) oPay.Invoices.TotalDiscountFC = pay.AppliedInvoices[i].TotalDiscountFC;

                        foreach (PropertyInfo info in pay.AppliedInvoices[i].UserFields.GetType().GetProperties())
                        {
                            if (pay.AppliedInvoices[i].UserFields.GetType().GetProperty(info.Name).GetValue(pay.AppliedInvoices[i].UserFields) != null)
                            {
                                oPay.Invoices.UserFields.Fields.Item(info.Name).Value = pay.AppliedInvoices[i].UserFields.GetType().GetProperty(info.Name).GetValue(pay.AppliedInvoices[i].UserFields);
                            }
                        }
                    }
                }
                #endregion
                #region Payment to account
                if (pay.AccountsPayment != null)
                {
                    for (int i = 0; i < pay.AccountsPayment.Count; i++)
                    {
                        if (i > 0) oPay.AccountPayments.Add();
                        oPay.AccountPayments.SetCurrentLine(i);
                        oPay.AccountPayments.AccountCode = pay.AccountsPayment[i].AccountCode;
                        if (String.IsNullOrEmpty(pay.AccountsPayment[i].Decription)) oPay.AccountPayments.Decription = pay.AccountsPayment[i].Decription;
                        if (pay.AccountsPayment[i].GrossAmount != 0) oPay.AccountPayments.GrossAmount = pay.AccountsPayment[i].GrossAmount;
                        if (!String.IsNullOrEmpty(pay.AccountsPayment[i].ProfitCenter)) oPay.AccountPayments.ProfitCenter = pay.AccountsPayment[i].ProfitCenter;
                        if (!String.IsNullOrEmpty(pay.AccountsPayment[i].ProfitCenter2)) oPay.AccountPayments.ProfitCenter2 = pay.AccountsPayment[i].ProfitCenter2;
                        if (!String.IsNullOrEmpty(pay.AccountsPayment[i].ProfitCenter3)) oPay.AccountPayments.ProfitCenter3 = pay.AccountsPayment[i].ProfitCenter3;
                        if (!String.IsNullOrEmpty(pay.AccountsPayment[i].ProfitCenter4)) oPay.AccountPayments.ProfitCenter4 = pay.AccountsPayment[i].ProfitCenter4;
                        if (!String.IsNullOrEmpty(pay.AccountsPayment[i].ProfitCenter5)) oPay.AccountPayments.ProfitCenter5 = pay.AccountsPayment[i].ProfitCenter5;
                        if (!String.IsNullOrEmpty(pay.AccountsPayment[i].ProjectCode)) oPay.AccountPayments.ProjectCode = pay.AccountsPayment[i].ProjectCode;
                        if (pay.AccountsPayment[i].SumPaid != 0) oPay.AccountPayments.SumPaid = pay.AccountsPayment[i].SumPaid;
                        if (!String.IsNullOrEmpty(pay.AccountsPayment[i].VatGroup)) oPay.AccountPayments.VatGroup = pay.AccountsPayment[i].VatGroup;
                        foreach (PropertyInfo info in pay.AccountsPayment[i].UserFields.GetType().GetProperties())
                        {
                            if (pay.AccountsPayment[i].UserFields.GetType().GetProperty(info.Name).GetValue(pay.AccountsPayment[i].UserFields) != null)
                            {
                                oPay.AccountPayments.UserFields.Fields.Item(info.Name).Value = pay.AccountsPayment[i].UserFields.GetType().GetProperty(info.Name).GetValue(pay.AccountsPayment[i].UserFields);
                            }
                        }
                    }
                }
                #endregion

                #region cheque payment
                if (pay.ChequesPayment != null)
                {
                    for (int i = 0; i < pay.ChequesPayment.Count; i++)
                    {
                        if (i > 0) oPay.Checks.Add();
                        oPay.Checks.SetCurrentLine(i);
                        if (!String.IsNullOrEmpty(pay.ChequesPayment[i].AccounttNum)) oPay.Checks.AccounttNum = pay.ChequesPayment[i].AccounttNum;
                        if (!String.IsNullOrEmpty(pay.ChequesPayment[i].BankCode)) oPay.Checks.BankCode = pay.ChequesPayment[i].BankCode;
                        if (!String.IsNullOrEmpty(pay.ChequesPayment[i].Branch)) oPay.Checks.Branch = pay.ChequesPayment[i].Branch;
                        if (!String.IsNullOrEmpty(pay.ChequesPayment[i].CheckAccount)) oPay.Checks.CheckAccount = pay.ChequesPayment[i].CheckAccount;
                        if (pay.ChequesPayment[i].CheckNumber != 0) oPay.Checks.CheckNumber = pay.ChequesPayment[i].CheckNumber;
                        if (pay.ChequesPayment[i].CheckSum != 0) oPay.Checks.CheckSum = pay.ChequesPayment[i].CheckSum;
                        if (!String.IsNullOrEmpty(pay.ChequesPayment[i].CountryCode)) oPay.Checks.CountryCode = pay.ChequesPayment[i].CountryCode;
                        if (!String.IsNullOrEmpty(pay.ChequesPayment[i].Details)) oPay.Checks.Details = pay.ChequesPayment[i].Details;
                        if (pay.ChequesPayment[i].DueDate != DateTime.MinValue) oPay.Checks.DueDate = pay.ChequesPayment[i].DueDate;
                        if (pay.ChequesPayment[i].ManualCheck) oPay.Checks.ManualCheck = SAPbobsCOM.BoYesNoEnum.tYES;
                        if (pay.ChequesPayment[i].Trnsfrable) oPay.Checks.Trnsfrable = SAPbobsCOM.BoYesNoEnum.tYES;

                        foreach (PropertyInfo info in pay.ChequesPayment[i].UserFields.GetType().GetProperties())
                        {
                            if (pay.ChequesPayment[i].UserFields.GetType().GetProperty(info.Name).GetValue(pay.ChequesPayment[i].UserFields) != null)
                            {
                                oPay.Checks.UserFields.Fields.Item(info.Name).Value = pay.ChequesPayment[i].UserFields.GetType().GetProperty(info.Name).GetValue(pay.ChequesPayment[i].UserFields);
                            }
                        }
                    }
                }
                #endregion

                #region Primary Form Items
                if (pay.FormItems != null)
                {
                    for (int i = 0; i < pay.FormItems.Count; i++)
                    {
                        if (i > 0) oPay.PrimaryFormItems.Add();
                        oPay.PrimaryFormItems.SetCurrentLine(i);
                        if (pay.FormItems[i].AmountFC != 0) oPay.PrimaryFormItems.AmountFC = pay.FormItems[i].AmountFC;
                        if (pay.FormItems[i].AmountLC != 0) oPay.PrimaryFormItems.AmountLC = pay.FormItems[i].AmountLC;
                        if (pay.FormItems[i].CashFlowLineItemID != 0) oPay.PrimaryFormItems.CashFlowLineItemID = pay.FormItems[i].CashFlowLineItemID;
                        if (!String.IsNullOrEmpty(pay.FormItems[i].CheckNumber)) oPay.PrimaryFormItems.CheckNumber = pay.FormItems[i].CheckNumber;
                        if (pay.FormItems[i].PaymentMeans != 0) oPay.PrimaryFormItems.PaymentMeans = (SAPbobsCOM.PaymentMeansTypeEnum)pay.FormItems[i].PaymentMeans;
                    }
                }
                #endregion

                if (oPay.Add() != 0)
                {
                    errMsg = oCom.GetLastErrorDescription();
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oPay = null;
                    return false;
                }
                oCom.GetNewObjectCode(out docentry);
                if (docentry == "")
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oPay = null;
                    return false;
                }
                SAPbobsCOM.Payments getDoc = (SAPbobsCOM.Payments)oCom.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), pay.DocObject));

                if (!getDoc.GetByKey(int.Parse(docentry)))
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    pay = null;
                    return false;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oPay);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(getDoc);
                oPay = null;
                getDoc = null;

                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

            }
            catch (Exception ex)
            {
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                errMsg = ex.Message;
                return false;
            }
            return true;
        }

        public bool CreateItem(Models.SAPModels.Items item, ref string key)
        {
            if (!generateLog(item, "4", item.LogUserID)) return false;
            oCom.StartTransaction();
            try
            {
                SAPbobsCOM.Items oItem = (SAPbobsCOM.Items)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
                #region Header 
                if (!String.IsNullOrEmpty(item.ItemCode)) oItem.ItemCode = item.ItemCode;
                if (!String.IsNullOrEmpty(item.ItemName)) oItem.ItemName = item.ItemName;
                if (String.IsNullOrEmpty(item.ApTaxCode)) oItem.ApTaxCode = item.ApTaxCode;
                if (!String.IsNullOrEmpty(item.ArTaxCode)) oItem.ArTaxCode = item.ArTaxCode;
                if (!String.IsNullOrEmpty(item.AssetClass)) oItem.AssetClass = item.AssetClass;
                if (!String.IsNullOrEmpty(item.AssetGroup)) oItem.AssetGroup = item.AssetGroup;
                if (item.AssetItem) oItem.AssetItem = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(item.AssetSerialNumber)) oItem.AssetSerialNumber = item.AssetSerialNumber;
                if (item.AutoCreateSerialNumbersOnRelease) oItem.AutoCreateSerialNumbersOnRelease = SAPbobsCOM.BoYesNoEnum.tYES;
                if (item.AvgStdPrice != 0) oItem.AvgStdPrice = item.AvgStdPrice;
                if (!String.IsNullOrEmpty(item.BarCode)) oItem.BarCode = item.BarCode;
                if (!String.IsNullOrEmpty(item.BaseUnitName)) oItem.BaseUnitName = item.BaseUnitName;
                if (item.BeverageCommercialBrandCode != 0) oItem.BeverageCommercialBrandCode = item.BeverageCommercialBrandCode;
                if (!String.IsNullOrEmpty(item.BeverageGroupCode)) oItem.BeverageGroupCode = item.BeverageGroupCode;
                if (!String.IsNullOrEmpty(item.BeverageTableCode)) oItem.BeverageTableCode = item.BeverageTableCode;
                if (item.CapitalizationDate != DateTime.MinValue) oItem.CapitalizationDate = item.CapitalizationDate;
                if (item.Cession) oItem.Cession = SAPbobsCOM.BoYesNoEnum.tYES;
                if (item.ChapterID != 0) oItem.ChapterID = item.ChapterID;
                if (item.CommissionGroup != 0) oItem.CommissionGroup = item.CommissionGroup;
                if (item.CommissionPercent != 0) oItem.CommissionPercent = item.CommissionPercent;
                if (item.CommissionSum != 0) oItem.CommissionSum = item.CommissionSum;
                if (item.CostAccountingMethod > 0) oItem.CostAccountingMethod = (SAPbobsCOM.BoInventorySystem)item.CostAccountingMethod;
                if (item.CustomsGroupCode != 0) oItem.CustomsGroupCode = item.CustomsGroupCode;
                if (!String.IsNullOrEmpty(item.DataExportCode)) oItem.DataExportCode = item.DataExportCode;
                if (item.DeactivateAfterUsefulLife) oItem.DeactivateAfterUsefulLife = SAPbobsCOM.BoYesNoEnum.tYES;
                if (item.DefaultPurchasingUoMEntry != 0) oItem.DefaultPurchasingUoMEntry = item.DefaultPurchasingUoMEntry;
                if (item.DefaultSalesUoMEntry != 0) oItem.DefaultSalesUoMEntry = item.DefaultSalesUoMEntry;
                if (!String.IsNullOrEmpty(item.DefaultWarehouse)) oItem.DefaultWarehouse = item.DefaultWarehouse;
                if (!String.IsNullOrEmpty(item.DepreciationGroup)) oItem.DepreciationGroup = item.DepreciationGroup;
                if (item.DesiredInventory != 0) oItem.DesiredInventory = item.DesiredInventory;
                if (!String.IsNullOrEmpty(item.ECExpensesAccount)) oItem.ECExpensesAccount = item.ECExpensesAccount;
                if (!String.IsNullOrEmpty(item.ECRevenuesAccount)) oItem.ECRevenuesAccount = item.ECRevenuesAccount;
                if (item.Employee != 0) oItem.Employee = item.Employee;
                if (item.ForceSelectionOfSerialNumber) oItem.ForceSelectionOfSerialNumber = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(item.ForeignExpensesAccount)) oItem.ForeignExpensesAccount = item.ForeignExpensesAccount;
                if (!String.IsNullOrEmpty(item.ForeignName)) oItem.ForeignName = item.ForeignName;
                if (!String.IsNullOrEmpty(item.ForeignRevenuesAccount)) oItem.ForeignRevenuesAccount = item.ForeignRevenuesAccount;
                if (item.Frozen) oItem.Frozen = SAPbobsCOM.BoYesNoEnum.tYES;
                if (item.FrozenFrom != DateTime.MinValue) oItem.FrozenFrom = item.FrozenFrom;
                if (!String.IsNullOrEmpty(item.FrozenRemarks)) oItem.FrozenRemarks = item.FrozenRemarks;
                if (item.FrozenTo != DateTime.MinValue) oItem.FrozenTo = item.FrozenTo;
                if (item.FuelID != 0) oItem.FuelID = item.FuelID;
                oItem.GLMethod = (SAPbobsCOM.BoGLMethods)item.GLMethod;
                if (!String.IsNullOrEmpty(item.GTSItemSpec)) oItem.GTSItemSpec = item.GTSItemSpec;
                if (!String.IsNullOrEmpty(item.GTSItemTaxCategory)) oItem.GTSItemTaxCategory = item.GTSItemTaxCategory;
                if (!String.IsNullOrEmpty(item.IncomeAccount)) oItem.IncomeAccount = item.IncomeAccount;
                //if (item.IncomingServiceCode > 0) oItem.IncomingServiceCode = item.IncomingServiceCode; //cluster B only
                if (item.IndirectTax) oItem.IndirectTax = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!item.InventoryItem) oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO;
                if (!String.IsNullOrEmpty(item.InventoryNumber)) oItem.InventoryNumber = item.InventoryNumber;
                if (!String.IsNullOrEmpty(item.InventoryUOM)) oItem.InventoryUOM = item.InventoryUOM;
                if (item.InventoryUoMEntry != 0) oItem.InventoryUoMEntry = item.InventoryUoMEntry;
                if (item.IsPhantom) oItem.IsPhantom = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(item.ItemCountryOrg)) oItem.ItemCountryOrg = item.ItemCountryOrg;
                if (item.IssueMethod > 0) oItem.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Manual; else oItem.IssueMethod = SAPbobsCOM.BoIssueMethod.im_Backflush;
                //if (item.IssuePrimarilyBy > 0) oItem.IssuePrimarilyBy = SAPbobsCOM.IssuePrimarilyByEnum.ipbBinLocations; else oItem.IssuePrimarilyBy = SAPbobsCOM.IssuePrimarilyByEnum.ipbSerialAndBatchNumbers;
                if (item.ItemsGroupCode != 0) oItem.ItemsGroupCode = item.ItemsGroupCode;
                if (item.LeadTime > 0) oItem.LeadTime = item.LeadTime;
                if (item.Location > 0) oItem.Location = item.Location;
                oItem.ItemType = (SAPbobsCOM.ItemTypeEnum)item.ItemType;
                if (!String.IsNullOrEmpty(item.Mainsupplier)) oItem.Mainsupplier = item.Mainsupplier;
                if (item.ManageBatchNumbers) oItem.ManageBatchNumbers = SAPbobsCOM.BoYesNoEnum.tYES;
                if (item.ManageSerialNumbers) oItem.ManageSerialNumbers = SAPbobsCOM.BoYesNoEnum.tYES;
                if (item.ManageSerialNumbersOnReleaseOnly) oItem.ManageSerialNumbersOnReleaseOnly = SAPbobsCOM.BoYesNoEnum.tYES;
                if (item.ManageStockByWarehouse) oItem.ManageStockByWarehouse = SAPbobsCOM.BoYesNoEnum.tYES; else oItem.ManageStockByWarehouse = SAPbobsCOM.BoYesNoEnum.tNO;
                if (item.Manufacturer != 0) oItem.Manufacturer = item.Manufacturer;
                if (item.MaterialGroup != 0) oItem.MaterialGroup = item.MaterialGroup;
                if (item.MaxInventory != 0) oItem.MaxInventory = item.MaxInventory;
                if (item.MinInventory != 0) oItem.MinInventory = item.MinInventory;
                if (item.MinOrderQuantity != 0) oItem.MinOrderQuantity = item.MinOrderQuantity;
                if (item.NoDiscounts) oItem.NoDiscounts = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(item.OrderIntervals)) oItem.OrderIntervals = item.OrderIntervals;
                if (item.OrderMultiple != 0) oItem.OrderMultiple = item.OrderMultiple;
                if (item.OutgoingServiceCode != 0) oItem.OutgoingServiceCode = item.OutgoingServiceCode;
                if (!String.IsNullOrEmpty(item.Picture)) oItem.Picture = item.Picture;
                oItem.PlanningSystem = (SAPbobsCOM.BoPlanningSystem)item.PlanningSystem;
                oItem.ProcurementMethod = (SAPbobsCOM.BoProcurementMethod)item.ProcurementMethod;
                if (item.ProductSource > 0) oItem.ProductSource = item.ProductSource;
                if (item.PurchaseFactor1 != 0) oItem.PurchaseFactor1 = item.PurchaseFactor1;
                if (item.PurchaseFactor2 != 0) oItem.PurchaseFactor2 = item.PurchaseFactor2;
                if (item.PurchaseFactor3 != 0) oItem.PurchaseFactor3 = item.PurchaseFactor3;
                if (item.PurchaseFactor4 != 0) oItem.PurchaseFactor4 = item.PurchaseFactor4;
                if (item.PurchaseHeightUnit != 0) oItem.PurchaseHeightUnit = item.PurchaseHeightUnit;
                if (item.PurchaseHeightUnit1 != 0) oItem.PurchaseHeightUnit1 = item.PurchaseHeightUnit1;
                if (!item.PurchaseItem) oItem.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tNO;
                if (item.PurchaseItemsPerUnit != 0) oItem.PurchaseItemsPerUnit = item.PurchaseItemsPerUnit;
                if (item.PurchaseLengthUnit != 0) oItem.PurchaseLengthUnit = item.PurchaseLengthUnit;
                if (item.PurchaseLengthUnit1 != 0) oItem.PurchaseLengthUnit1 = item.PurchaseLengthUnit1;
                if (!String.IsNullOrEmpty(item.PurchasePackagingUnit)) oItem.PurchasePackagingUnit = item.PurchasePackagingUnit;
                if (item.PurchaseQtyPerPackUnit != 0) oItem.PurchaseQtyPerPackUnit = item.PurchaseQtyPerPackUnit;
                if (!String.IsNullOrEmpty(item.PurchaseUnit)) oItem.PurchaseUnit = item.PurchaseUnit;
                if (item.PurchaseUnitHeight != 0) oItem.PurchaseUnitHeight = item.PurchaseUnitHeight;
                if (item.PurchaseUnitHeight1 != 0) oItem.PurchaseUnitHeight1 = item.PurchaseUnitHeight1;
                if (item.PurchaseUnitLength != 0) oItem.PurchaseUnitLength = item.PurchaseUnitLength;
                if (item.PurchaseUnitLength1 != 0) oItem.PurchaseUnitLength1 = item.PurchaseUnitLength1;
                if (item.PurchaseUnitVolume != 0) oItem.PurchaseUnitVolume = item.PurchaseUnitVolume;
                if (item.PurchaseUnitWeight != 0) oItem.PurchaseUnitWeight = item.PurchaseUnitWeight;
                if (item.PurchaseUnitWeight1 != 0) oItem.PurchaseUnitWeight1 = item.PurchaseUnitWeight1;
                if (item.PurchaseUnitWidth != 0) oItem.PurchaseUnitWidth = item.PurchaseUnitWidth;
                if (item.PurchaseUnitWidth1 != 0) oItem.PurchaseUnitWidth1 = item.PurchaseUnitWidth1;
                if (!String.IsNullOrEmpty(item.PurchaseVATGroup)) oItem.PurchaseVATGroup = item.PurchaseVATGroup;
                if (item.PurchaseVolumeUnit != 0) oItem.PurchaseVolumeUnit = item.PurchaseVolumeUnit;
                if (item.PurchaseWeightUnit != 0) oItem.PurchaseWeightUnit = item.PurchaseWeightUnit;
                if (item.PurchaseWeightUnit1 != 0) oItem.PurchaseWeightUnit1 = item.PurchaseWeightUnit1;
                if (item.PurchaseWidthUnit != 0) oItem.PurchaseWidthUnit = item.PurchaseWidthUnit;
                if (item.PurchaseWidthUnit1 != 0) oItem.PurchaseWidthUnit1 = item.PurchaseWidthUnit1;
                if (item.SalesFactor1 != 0) oItem.SalesFactor1 = item.SalesFactor1;
                if (item.SalesFactor2 != 0) oItem.SalesFactor2 = item.SalesFactor2;
                if (item.SalesFactor3 != 0) oItem.SalesFactor3 = item.SalesFactor3;
                if (item.SalesFactor4 != 0) oItem.SalesFactor4 = item.SalesFactor4;
                if (item.SalesHeightUnit != 0) oItem.SalesHeightUnit = item.SalesHeightUnit;
                if (item.SalesHeightUnit1 != 0) oItem.SalesHeightUnit1 = item.SalesHeightUnit1;
                if (!item.SalesItem) oItem.SalesItem = SAPbobsCOM.BoYesNoEnum.tNO;
                if (item.SalesItemsPerUnit != 0) oItem.SalesItemsPerUnit = item.SalesItemsPerUnit;
                if (item.SalesLengthUnit != 0) oItem.SalesLengthUnit = item.SalesLengthUnit;
                if (item.SalesLengthUnit1 != 0) oItem.SalesLengthUnit1 = item.SalesLengthUnit1;
                if (!String.IsNullOrEmpty(item.SalesPackagingUnit)) oItem.SalesPackagingUnit = oItem.SalesPackagingUnit;
                if (item.SalesQtyPerPackUnit != 0) oItem.SalesQtyPerPackUnit = item.SalesQtyPerPackUnit;
                if (!String.IsNullOrEmpty(item.SalesUnit)) oItem.SalesUnit = item.SalesUnit;
                if (item.SalesUnitHeight != 0) oItem.SalesUnitHeight = item.SalesUnitHeight;
                if (item.SalesUnitHeight1 != 0) oItem.SalesUnitHeight1 = item.SalesUnitHeight1;
                if (item.SalesUnitLength != 0) oItem.SalesUnitLength = item.SalesUnitLength;
                if (item.SalesUnitLength1 != 0) oItem.SalesUnitLength1 = item.SalesUnitLength1;
                if (item.SalesUnitVolume != 0) oItem.SalesUnitVolume = item.SalesUnitVolume;
                if (item.SalesUnitWeight != 0) oItem.SalesUnitWeight = item.SalesUnitWeight;
                if (item.SalesUnitWeight1 != 0) oItem.SalesUnitWeight1 = item.SalesUnitWeight1;
                if (item.SalesUnitWidth != 0) oItem.SalesUnitWidth = item.SalesUnitWidth;
                if (item.SalesUnitWidth1 != 0) oItem.SalesUnitWidth1 = item.SalesUnitWidth1;
                if (!String.IsNullOrEmpty(item.SalesVATGroup)) oItem.SalesVATGroup = item.SalesVATGroup;
                if (item.SalesVolumeUnit != 0) oItem.SalesVolumeUnit = item.SalesVolumeUnit;
                if (item.SalesWeightUnit != 0) oItem.SalesWeightUnit = item.SalesWeightUnit;
                if (item.SalesWeightUnit1 != 0) oItem.SalesWeightUnit1 = item.SalesWeightUnit1;
                if (item.SalesWidthUnit != 0) oItem.SalesWidthUnit = item.SalesWidthUnit;
                if (item.SalesWidthUnit1 != 0) oItem.SalesWidthUnit1 = item.SalesWidthUnit1;
                if (!String.IsNullOrEmpty(item.ScsCode)) oItem.ScsCode = item.ScsCode;
                if (!String.IsNullOrEmpty(item.SerialNum)) oItem.SerialNum = item.SerialNum;
                if (item.Series > 0) oItem.Series = item.Series;
                if (item.ServiceGroup != 0) oItem.ServiceGroup = item.ServiceGroup;
                if (item.ShipType != 0) oItem.ShipType = item.ShipType;
                if (item.StatisticalAsset) oItem.StatisticalAsset = SAPbobsCOM.BoYesNoEnum.tYES;
                if (item.SRIAndBatchManageMethod != 0) oItem.SRIAndBatchManageMethod = (SAPbobsCOM.BoManageMethod)oItem.SRIAndBatchManageMethod;
                if (!String.IsNullOrEmpty(item.SupplierCatalogNo)) oItem.SupplierCatalogNo = item.SupplierCatalogNo;
                if (!String.IsNullOrEmpty(item.SWW)) oItem.SWW = item.SWW;
                if (item.Technician != 0) oItem.Technician = item.Technician;
                if (item.ToleranceDays != 0) oItem.ToleranceDays = item.ToleranceDays;
                if (item.UoMGroupEntry != 0) oItem.UoMGroupEntry = item.UoMGroupEntry;
                if (item.Valid) oItem.Valid = SAPbobsCOM.BoYesNoEnum.tYES;
                if (item.ValidFrom != DateTime.MinValue) oItem.ValidFrom = item.ValidFrom;
                if (!String.IsNullOrEmpty(item.ValidRemarks)) oItem.ValidRemarks = item.ValidRemarks;
                if (item.ValidTo != DateTime.MinValue) oItem.ValidTo = item.ValidTo;
                if (item.VatLiable) oItem.VatLiable = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(item.WarrantyTemplate)) oItem.WarrantyTemplate = item.WarrantyTemplate;
                if (item.WTLiable) oItem.WTLiable = SAPbobsCOM.BoYesNoEnum.tYES;

                foreach (PropertyInfo info in item.UserFields.GetType().GetProperties())
                {
                    if (item.UserFields.GetType().GetProperty(info.Name).GetValue(item.UserFields) != null)
                    {
                        oItem.UserFields.Fields.Item(info.Name).Value = item.UserFields.GetType().GetProperty(info.Name).GetValue(item.UserFields);
                    }
                }
                #endregion 
                #region Warehouse
                if (item.WhsInfo != null)
                {
                    for (int i = 0; i < item.WhsInfo.Count; i++)
                    {
                        if (i > 0) oItem.WhsInfo.Add();
                        oItem.WhsInfo.SetCurrentLine(i);
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].WarehouseCode)) oItem.WhsInfo.WarehouseCode = item.WhsInfo[i].WarehouseCode;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].CostAccount)) oItem.WhsInfo.CostAccount = item.WhsInfo[i].CostAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].CostInflationAccount)) oItem.WhsInfo.CostInflationAccount = item.WhsInfo[i].CostInflationAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].CostInflationOffsetAccount)) oItem.WhsInfo.CostInflationOffsetAccount = item.WhsInfo[i].CostInflationOffsetAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].DecreasingAccount)) oItem.WhsInfo.DecreasingAccount = item.WhsInfo[i].DecreasingAccount;
                        if (item.WhsInfo[i].DefaultBin != 0) oItem.WhsInfo.DefaultBin = item.WhsInfo[i].DefaultBin;
                        if (item.WhsInfo[i].DefaultBinEnforced) oItem.WhsInfo.DefaultBinEnforced = SAPbobsCOM.BoYesNoEnum.tYES;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].EUExpensesAccount)) oItem.WhsInfo.EUExpensesAccount = item.WhsInfo[i].EUExpensesAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].EUPurchaseCreditAcc)) oItem.WhsInfo.EUPurchaseCreditAcc = item.WhsInfo[i].EUPurchaseCreditAcc;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].EURevenuesAccount)) oItem.WhsInfo.EURevenuesAccount = item.WhsInfo[i].EURevenuesAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ExchangeRateDifferencesAcct)) oItem.WhsInfo.ExchangeRateDifferencesAcct = item.WhsInfo[i].ExchangeRateDifferencesAcct;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ExemptedCredits)) oItem.WhsInfo.ExemptedCredits = item.WhsInfo[i].ExemptedCredits;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ExemptIncomeAcc)) oItem.WhsInfo.ExemptIncomeAcc = item.WhsInfo[i].ExemptIncomeAcc;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ExpenseClearingAct)) oItem.WhsInfo.ExpenseClearingAct = item.WhsInfo[i].ExpenseClearingAct;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ExpenseOffsettingAccount)) oItem.WhsInfo.ExpenseOffsettingAccount = item.WhsInfo[i].ExpenseOffsettingAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ExpensesAccount)) oItem.WhsInfo.ExpensesAccount = item.WhsInfo[i].ExpensesAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ForeignExpensAcc)) oItem.WhsInfo.ForeignExpensAcc = item.WhsInfo[i].ForeignExpensAcc;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ForeignPurchaseCreditAcc)) oItem.WhsInfo.ForeignPurchaseCreditAcc = item.WhsInfo[i].ForeignPurchaseCreditAcc;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ForeignRevenueAcc)) oItem.WhsInfo.ForeignRevenueAcc = item.WhsInfo[i].ForeignRevenueAcc;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].GLDecreaseAcct)) oItem.WhsInfo.GLDecreaseAcct = item.WhsInfo[i].GLDecreaseAcct;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].GLIncreaseAcct)) oItem.WhsInfo.GLIncreaseAcct = item.WhsInfo[i].GLIncreaseAcct;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].GoodsClearingAcct)) oItem.WhsInfo.GoodsClearingAcct = item.WhsInfo[i].GoodsClearingAcct;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].IncreasingAccount)) oItem.WhsInfo.IncreasingAccount = item.WhsInfo[i].IncreasingAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].InventoryAccount)) oItem.WhsInfo.InventoryAccount = item.WhsInfo[i].InventoryAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].InventoryOffsetProfitAndLossAccount)) oItem.WhsInfo.InventoryOffsetProfitAndLossAccount = item.WhsInfo[i].InventoryOffsetProfitAndLossAccount;
                        if (item.WhsInfo[i].Locked) oItem.WhsInfo.Locked = SAPbobsCOM.BoYesNoEnum.tYES;
                        if (item.WhsInfo[i].MaximalStock != 0) oItem.WhsInfo.MaximalStock = item.WhsInfo[i].MaximalStock;
                        if (item.WhsInfo[i].MinimalOrder != 0) oItem.WhsInfo.MinimalOrder = item.WhsInfo[i].MinimalOrder;
                        if (item.WhsInfo[i].MinimalStock != 0) oItem.WhsInfo.MinimalStock = item.WhsInfo[i].MinimalStock;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].NegativeInventoryAdjustmentAccount)) oItem.WhsInfo.NegativeInventoryAdjustmentAccount = item.WhsInfo[i].NegativeInventoryAdjustmentAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].PAReturnAcct)) oItem.WhsInfo.PAReturnAcct = item.WhsInfo[i].PAReturnAcct;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].PriceDifferenceAcc)) oItem.WhsInfo.PriceDifferenceAcc = item.WhsInfo[i].PriceDifferenceAcc;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].PurchaseAcct)) oItem.WhsInfo.PurchaseAcct = item.WhsInfo[i].PurchaseAcct;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].PurchaseBalanceAccount)) oItem.WhsInfo.PurchaseBalanceAccount = item.WhsInfo[i].PurchaseBalanceAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].PurchaseCreditAcc)) oItem.WhsInfo.PurchaseCreditAcc = item.WhsInfo[i].PurchaseCreditAcc;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].PurchaseOffsetAcct)) oItem.WhsInfo.PurchaseOffsetAcct = item.WhsInfo[i].PurchaseOffsetAcct;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ReturningAccount)) oItem.WhsInfo.ReturningAccount = item.WhsInfo[i].ReturningAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].RevenuesAccount)) oItem.WhsInfo.RevenuesAccount = item.WhsInfo[i].RevenuesAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].SalesCreditAcc)) oItem.WhsInfo.SalesCreditAcc = item.WhsInfo[i].SalesCreditAcc;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].SalesCreditEUAcc)) oItem.WhsInfo.SalesCreditEUAcc = item.WhsInfo[i].SalesCreditEUAcc;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].SalesCreditForeignAcc)) oItem.WhsInfo.SalesCreditForeignAcc = item.WhsInfo[i].SalesCreditForeignAcc;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].ShippedGoodsAccount)) oItem.WhsInfo.ShippedGoodsAccount = item.WhsInfo[i].ShippedGoodsAccount;
                        if (item.WhsInfo[i].StandardAveragePrice != 0) oItem.WhsInfo.StandardAveragePrice = item.WhsInfo[i].StandardAveragePrice;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].StockInflationAdjustAccount)) oItem.WhsInfo.StockInflationAdjustAccount = item.WhsInfo[i].StockInflationAdjustAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].StockInflationOffsetAccount)) oItem.WhsInfo.StockInflationOffsetAccount = item.WhsInfo[i].StockInflationOffsetAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].StockInTransitAccount)) oItem.WhsInfo.StockInTransitAccount = item.WhsInfo[i].StockInTransitAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].TransferAccount)) oItem.WhsInfo.TransferAccount = item.WhsInfo[i].TransferAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].VarienceAccount)) oItem.WhsInfo.VarienceAccount = item.WhsInfo[i].VarienceAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].VATInRevenueAccount)) oItem.WhsInfo.VATInRevenueAccount = item.WhsInfo[i].VATInRevenueAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].WHIncomingCenvatAccount)) oItem.WhsInfo.WHIncomingCenvatAccount = item.WhsInfo[i].WHIncomingCenvatAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].WHOutgoingCenvatAccount)) oItem.WhsInfo.WHOutgoingCenvatAccount = item.WhsInfo[i].WHOutgoingCenvatAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].WipAccount)) oItem.WhsInfo.WipAccount = item.WhsInfo[i].WipAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].WipOffsetProfitAndLossAccount)) oItem.WhsInfo.WipOffsetProfitAndLossAccount = item.WhsInfo[i].WipOffsetProfitAndLossAccount;
                        if (!String.IsNullOrEmpty(item.WhsInfo[i].WipVarianceAccount)) oItem.WhsInfo.WipVarianceAccount = item.WhsInfo[i].WipVarianceAccount;
                        foreach (PropertyInfo info in item.WhsInfo[i].UserFields.GetType().GetProperties())
                        {
                            if (item.WhsInfo[i].UserFields.GetType().GetProperty(info.Name).GetValue(item.WhsInfo[i].UserFields) != null)
                            {
                                oItem.WhsInfo.UserFields.Fields.Item(info.Name).Value = item.WhsInfo[i].UserFields.GetType().GetProperty(info.Name).GetValue(item.WhsInfo[i].UserFields);
                            }
                        }
                    }
                }
                #endregion
                #region Vendor
                if (item.PreferedVendor != null)
                {
                    for (int i = 0; i < item.PreferedVendor.Count; i++)
                    {
                        if (i > 0) oItem.PreferredVendors.Add();
                        oItem.PreferredVendors.SetCurrentLine(i);
                        oItem.PreferredVendors.BPCode = item.PreferedVendor[i].BPCode;

                    }
                }
                #endregion

                if (oItem.Add() != 0)
                {
                    errMsg = oCom.GetLastErrorDescription();
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oItem = null;
                    return false;
                }
                oCom.GetNewObjectCode(out key);
                if (key == "")
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oItem = null;
                    return false;
                }
                SAPbobsCOM.Items getDoc = (SAPbobsCOM.Items)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);

                if (!getDoc.GetByKey(key))
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oItem = null;
                    return false;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(getDoc);
                oItem = null;
                getDoc = null;

                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                errMsg = ex.Message;
                return false;

            }
            return true;
        }

        public bool CreateBP(Models.SAPModels.BPMaster bp, ref string key)
        {
            if (!generateLog(bp, "2", bp.LogUserID)) return false;
            oCom.StartTransaction();
            try
            {
                SAPbobsCOM.BusinessPartners oBP = (SAPbobsCOM.BusinessPartners)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
                #region Header

                if (!String.IsNullOrEmpty(bp.CardCode)) oBP.CardCode = bp.CardCode;
                if (!String.IsNullOrEmpty(bp.CardName)) oBP.CardName = bp.CardName;
                if (!String.IsNullOrEmpty(bp.CardForeignName)) oBP.CardForeignName = bp.CardForeignName;
                if (bp.AccrualCriteria) oBP.AccrualCriteria = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.AdditionalID)) oBP.AdditionalID = bp.AdditionalID;
                if (!String.IsNullOrEmpty(bp.Address)) oBP.Address = bp.Address;
                if (!String.IsNullOrEmpty(bp.AliasName)) oBP.AliasName = bp.AliasName;
                if (bp.Affiliate) oBP.Affiliate = SAPbobsCOM.BoYesNoEnum.tYES;
                if (bp.AvarageLate != 0) oBP.AvarageLate = bp.AvarageLate;
                if (bp.AutomaticPosting != 0) oBP.AutomaticPosting = (SAPbobsCOM.AutomaticPostingEnum)bp.AutomaticPosting;
                if (bp.BackOrder) oBP.BackOrder = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.BankChargesAllocationCode)) oBP.BankChargesAllocationCode = bp.BankChargesAllocationCode;
                if (!String.IsNullOrEmpty(bp.BankCountry)) oBP.BankCountry = bp.BankCountry;
                if (!String.IsNullOrEmpty(bp.BillofExchangeonCollection)) oBP.BillofExchangeonCollection = bp.BillofExchangeonCollection;
                if (!String.IsNullOrEmpty(bp.BillToBuildingFloorRoom)) oBP.BillToBuildingFloorRoom = bp.BillToBuildingFloorRoom;
                if (!String.IsNullOrEmpty(bp.BilltoDefault)) oBP.BilltoDefault = bp.BilltoDefault;
                if (!String.IsNullOrEmpty(bp.BillToState)) oBP.BillToState = bp.BillToState;
                if (!String.IsNullOrEmpty(bp.Block)) oBP.Block = bp.Block;
                if (bp.BlockDunning) oBP.BlockDunning = SAPbobsCOM.BoYesNoEnum.tYES;
                if (bp.BookkeepingCertified) oBP.BookkeepingCertified = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.Box1099)) oBP.Box1099 = bp.Box1099;
                if (!String.IsNullOrEmpty(bp.BusinessType)) oBP.BusinessType = bp.BusinessType;
                if (bp.CampaignNumber != 0) oBP.CampaignNumber = bp.CampaignNumber;
                oBP.CardType = (SAPbobsCOM.BoCardTypes)bp.CardType;
                if (String.IsNullOrEmpty(bp.Cellular)) oBP.Cellular = bp.Cellular;
                if (!String.IsNullOrEmpty(bp.CertificateNumber)) oBP.CertificateNumber = bp.CertificateNumber;
                if (!String.IsNullOrEmpty(bp.ChannelBP)) oBP.ChannelBP = bp.ChannelBP;
                if (!String.IsNullOrEmpty(bp.City)) oBP.City = bp.City;
                if (bp.CollectionAuthorization) oBP.CollectionAuthorization = SAPbobsCOM.BoYesNoEnum.tYES;
                if (bp.CommissionGroupCode != 0) oBP.CommissionGroupCode = bp.CommissionGroupCode;
                if (bp.CommissionPercent != 0) oBP.CommissionPercent = bp.CommissionPercent;
                oBP.CompanyPrivate = (SAPbobsCOM.BoCardCompanyTypes)bp.CompanyPrivate;
                if (!String.IsNullOrEmpty(bp.CompanyRegistrationNumber)) oBP.CompanyRegistrationNumber = bp.CompanyRegistrationNumber;
                if (!String.IsNullOrEmpty(bp.ContactPerson)) oBP.ContactPerson = bp.ContactPerson;
                if (!String.IsNullOrEmpty(bp.Country)) oBP.Country = bp.Country;
                if (!String.IsNullOrEmpty(bp.County)) oBP.County = bp.County;
                if (bp.CreditCardCode != 0) oBP.CreditCardCode = bp.CreditCardCode;
                if (bp.CreditCardExpiration != DateTime.MinValue) oBP.CreditCardExpiration = bp.CreditCardExpiration;
                if (!String.IsNullOrEmpty(bp.CreditCardNum)) oBP.CreditCardNum = bp.CreditCardNum;
                if (bp.CreditLimit != 0) oBP.CreditLimit = bp.CreditLimit;
                if (!String.IsNullOrEmpty(bp.Currency)) oBP.Currency = bp.Currency;
                if (!String.IsNullOrEmpty(bp.CustomerBillofExchangDisc)) oBP.CustomerBillofExchangDisc = bp.CustomerBillofExchangDisc;
                if (!String.IsNullOrEmpty(bp.CustomerBillofExchangPres)) oBP.CustomerBillofExchangPres = bp.CustomerBillofExchangPres;
                //if (bp.DatevAccount != 0) oBP.DatevAccount = bp.DatevAccount;
                if (bp.DatevFirstDataEntry) oBP.DatevFirstDataEntry = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.DebitorAccount)) oBP.DebitorAccount = bp.DebitorAccount;
                if (bp.DeductibleAtSource) oBP.DeductibleAtSource = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.DeductionOffice)) oBP.DeductionOffice = bp.DeductionOffice;
                if (bp.DeductionPercent != 0) oBP.DeductionPercent = bp.DeductionPercent;
                if (bp.DeductionValidUntil != DateTime.MinValue) oBP.DeductionValidUntil = bp.DeductionValidUntil;
                if (!String.IsNullOrEmpty(bp.DefaultAccount)) oBP.DefaultAccount = bp.DefaultAccount;
                if (!String.IsNullOrEmpty(bp.DefaultBankCode)) oBP.DefaultBankCode = bp.DefaultBankCode;
                if (bp.DefaultBlanketAgreementNumber != 0) oBP.DefaultBlanketAgreementNumber = bp.DefaultBlanketAgreementNumber;
                if (!String.IsNullOrEmpty(bp.DefaultBranch)) oBP.DefaultBranch = bp.DefaultBranch;
                if (bp.DefaultTechnician != 0) oBP.DefaultTechnician = bp.DefaultTechnician;
                if (bp.DeferredTax) oBP.DeferredTax = SAPbobsCOM.BoYesNoEnum.tYES;
                if (bp.DiscountBaseObject != 0) oBP.DiscountBaseObject = (SAPbobsCOM.DiscountGroupBaseObjectEnum)bp.DiscountBaseObject;
                if (bp.DiscountPercent != 0) oBP.DiscountPercent = bp.DiscountPercent;
                if (bp.DiscountRelations != 0) oBP.DiscountRelations = (SAPbobsCOM.DiscountGroupRelationsEnum)bp.DiscountRelations;
                if (!String.IsNullOrEmpty(bp.DME)) oBP.DME = bp.DME;
                if (!String.IsNullOrEmpty(bp.DownPaymentClearAct)) oBP.DownPaymentClearAct = bp.DownPaymentClearAct;
                if (!String.IsNullOrEmpty(bp.DownPaymentInterimAccount)) oBP.DownPaymentInterimAccount = bp.DownPaymentInterimAccount;
                if (!String.IsNullOrEmpty(bp.DunningTerm)) oBP.DunningTerm = bp.DunningTerm;
                if (!String.IsNullOrEmpty(bp.EDIRecipientID)) oBP.EDIRecipientID = bp.EDIRecipientID;
                if (!String.IsNullOrEmpty(bp.EDISenderID)) oBP.EDISenderID = bp.EDISenderID;
                if (bp.EffectiveDiscount != 0) oBP.EffectiveDiscount = (SAPbobsCOM.DiscountGroupRelationsEnum)bp.EffectiveDiscount;
                if (!String.IsNullOrEmpty(bp.EmailAddress)) oBP.EmailAddress = bp.EmailAddress;
                if (bp.Equalization) oBP.Equalization = SAPbobsCOM.BoYesNoEnum.tYES;
                if (bp.ETaxWebSite != 0) oBP.ETaxWebSite = bp.ETaxWebSite;
                if (bp.ExemptionValidityDateFrom != DateTime.MinValue) oBP.ExemptionValidityDateFrom = bp.ExemptionValidityDateFrom;
                if (bp.ExemptionValidityDateTo != DateTime.MinValue) oBP.ExemptionValidityDateTo = bp.ExemptionValidityDateTo;
                if (!String.IsNullOrEmpty(bp.ExemptNum)) oBP.ExemptNum = bp.ExemptNum;
                if (bp.ExpirationDate != DateTime.MinValue) oBP.ExpirationDate = bp.ExpirationDate;
                if (!String.IsNullOrEmpty(bp.ExportCode)) oBP.ExportCode = bp.ExportCode;
                if (!String.IsNullOrEmpty(bp.FatherCard)) { oBP.FatherCard = bp.FatherCard; oBP.FatherType = (SAPbobsCOM.BoFatherCardTypes)bp.FatherType; }
                if (!String.IsNullOrEmpty(bp.Fax)) oBP.Fax = bp.Fax;
                if (!String.IsNullOrEmpty(bp.FederalTaxID)) oBP.FederalTaxID = bp.FederalTaxID;
                if (!String.IsNullOrEmpty(bp.FeeAccount)) oBP.FeeAccount = bp.FeeAccount;
                if (bp.FormCode1099 != 0) oBP.FormCode1099 = bp.FormCode1099;
                if (!String.IsNullOrEmpty(bp.FreeText)) oBP.FreeText = bp.FreeText;
                if (bp.Frozen) oBP.Frozen = SAPbobsCOM.BoYesNoEnum.tYES;
                if (bp.FrozenFrom != DateTime.MinValue) oBP.FrozenFrom = bp.FrozenFrom;
                if (!String.IsNullOrEmpty(bp.FrozenRemarks)) oBP.FrozenRemarks = bp.FrozenRemarks;
                if (bp.FrozenTo != DateTime.MinValue) oBP.FrozenTo = bp.FrozenTo;
                if (!String.IsNullOrEmpty(bp.GlobalLocationNumber)) oBP.GlobalLocationNumber = bp.GlobalLocationNumber;
                if (bp.GroupCode != 0) oBP.GroupCode = bp.GroupCode;
                if (!String.IsNullOrEmpty(bp.GTSBankAccountNo)) oBP.GTSBankAccountNo = bp.GTSBankAccountNo;
                if (!String.IsNullOrEmpty(bp.GTSBillingAddrTel)) oBP.GTSBillingAddrTel = bp.GTSBillingAddrTel;
                if (!String.IsNullOrEmpty(bp.GTSRegNo)) oBP.GTSRegNo = bp.GTSRegNo;
                if (bp.HierarchicalDeduction) oBP.HierarchicalDeduction = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.HouseBank)) oBP.HouseBank = bp.HouseBank;
                if (!String.IsNullOrEmpty(bp.HouseBankAccount)) oBP.HouseBankAccount = bp.HouseBankAccount;
                if (!String.IsNullOrEmpty(bp.HouseBankBranch)) oBP.HouseBankBranch = bp.HouseBankBranch;
                if (!String.IsNullOrEmpty(bp.HouseBankCountry)) oBP.HouseBankCountry = bp.HouseBankCountry;
                if (!String.IsNullOrEmpty(bp.IBAN)) oBP.IBAN = bp.IBAN;
                if (!String.IsNullOrEmpty(bp.Indicator)) oBP.Indicator = bp.Indicator;
                if (bp.Industry != 0) oBP.Industry = bp.Industry;
                if (!String.IsNullOrEmpty(bp.IndustryType)) oBP.IndustryType = bp.IndustryType;
                if (!String.IsNullOrEmpty(bp.InstructionKey)) oBP.InstructionKey = bp.InstructionKey;
                if (bp.InsuranceOperation347) oBP.InsuranceOperation347 = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.InterestAccount)) oBP.InterestAccount = bp.InterestAccount;
                if (bp.IntrestRatePercent != 0) oBP.IntrestRatePercent = bp.IntrestRatePercent;
                if (!String.IsNullOrEmpty(bp.ISRBillerID)) oBP.ISRBillerID = bp.ISRBillerID;
                if (bp.LanguageCode != 0) oBP.LanguageCode = bp.LanguageCode;
                if (!String.IsNullOrEmpty(bp.LinkedBusinessPartner)) oBP.LinkedBusinessPartner = bp.LinkedBusinessPartner;
                if (!String.IsNullOrEmpty(bp.MailAddress)) oBP.MailAddress = bp.MailAddress;
                if (!String.IsNullOrEmpty(bp.MailCity)) oBP.MailCity = bp.MailCity;
                if (!String.IsNullOrEmpty(bp.MailCountry)) oBP.MailCountry = bp.MailCountry;
                if (!String.IsNullOrEmpty(bp.MailCounty)) oBP.MailCounty = bp.MailCounty;
                if (!String.IsNullOrEmpty(bp.MailZipCode)) oBP.MailZipCode = bp.MailZipCode;
                if (bp.MaxAmountOfExemption != 0) oBP.MaxAmountOfExemption = bp.MaxAmountOfExemption;
                if (bp.MaxCommitment != 0) oBP.MaxCommitment = bp.MaxCommitment;
                if (bp.MinIntrest != 0) oBP.MinIntrest = bp.MinIntrest;
                if (!String.IsNullOrEmpty(bp.NationalInsuranceNum)) oBP.NationalInsuranceNum = bp.NationalInsuranceNum;
                if (bp.NoDiscounts) oBP.NoDiscounts = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.Notes)) oBP.Notes = bp.Notes;
                if (!String.IsNullOrEmpty(bp.OtherReceivablePayable)) oBP.OtherReceivablePayable = bp.OtherReceivablePayable;
                if (!String.IsNullOrEmpty(bp.OwnerIDNumber)) oBP.OwnerIDNumber = bp.OwnerIDNumber;
                if (!String.IsNullOrEmpty(bp.Pager)) oBP.Pager = bp.Pager;
                if (bp.PartialDelivery) oBP.PartialDelivery = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.Password)) oBP.Password = bp.Password;
                if (bp.PaymentBlock) oBP.PaymentBlock = SAPbobsCOM.BoYesNoEnum.tYES;
                if (bp.PaymentBlockDescription != 0) oBP.PaymentBlockDescription = bp.PaymentBlockDescription;
                if (bp.PayTermsGrpCode != 0) oBP.PayTermsGrpCode = bp.PayTermsGrpCode;
                if (!String.IsNullOrEmpty(bp.PeymentMethodCode)) oBP.PeymentMethodCode = bp.PeymentMethodCode;
                if (!String.IsNullOrEmpty(bp.Phone1)) oBP.Phone1 = bp.Phone1;
                if (!String.IsNullOrEmpty(bp.Phone2)) oBP.Phone2 = bp.Phone2;
                if (!String.IsNullOrEmpty(bp.Picture)) oBP.Picture = bp.Picture;
                if (!String.IsNullOrEmpty(bp.PlanningGroup)) oBP.PlanningGroup = bp.PlanningGroup;
                if (bp.PriceListNum != 0) oBP.PriceListNum = bp.PriceListNum;
                if (bp.Priority != 0) oBP.Priority = bp.Priority;
                if (!String.IsNullOrEmpty(bp.Profession)) oBP.Profession = bp.Profession;
                if (!String.IsNullOrEmpty(bp.ProjectCode)) oBP.ProjectCode = bp.ProjectCode;
                if (!String.IsNullOrEmpty(bp.RateDiffAccount)) oBP.RateDiffAccount = bp.RateDiffAccount;
                if (!String.IsNullOrEmpty(bp.ReferenceDetails)) oBP.ReferenceDetails = bp.ReferenceDetails;
                if (!String.IsNullOrEmpty(bp.RelationshipCode)) oBP.RelationshipCode = bp.RelationshipCode;
                if (bp.RelationshipDateFrom != DateTime.MinValue) oBP.RelationshipDateFrom = bp.RelationshipDateFrom;
                if (bp.RelationshipDateTill != DateTime.MinValue) oBP.RelationshipDateTill = bp.RelationshipDateTill;
                if (!String.IsNullOrEmpty(bp.RepresentativeName)) oBP.RepresentativeName = bp.RepresentativeName;
                if (bp.SalesPersonCode != 0) oBP.SalesPersonCode = bp.SalesPersonCode;
                if (bp.Series != 0) oBP.Series = bp.Series;
                if (bp.ShippingType != 0) oBP.ShippingType = bp.ShippingType;
                if (!String.IsNullOrEmpty(bp.ShipToBuildingFloorRoom)) oBP.ShipToBuildingFloorRoom = bp.ShipToBuildingFloorRoom;
                if (!String.IsNullOrEmpty(bp.ShipToDefault)) oBP.ShipToDefault = bp.ShipToDefault;
                if (bp.SinglePayment) oBP.SinglePayment = SAPbobsCOM.BoYesNoEnum.tYES;
                if (bp.SubjectToWithholdingTax) oBP.SubjectToWithholdingTax = (BoYesNoNoneEnum)SAPbobsCOM.BoYesNoEnum.tYES;
                if (bp.SurchargeOverlook) oBP.SurchargeOverlook = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.TaxExemptionLetterNum)) oBP.TaxExemptionLetterNum = bp.TaxExemptionLetterNum;
                if (bp.Territory != 0) oBP.Territory = bp.Territory;
                if (bp.ThresholdOverlook) oBP.ThresholdOverlook = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(bp.UnifiedFederalTaxID)) oBP.UnifiedFederalTaxID = bp.UnifiedFederalTaxID;
                if (!String.IsNullOrEmpty(bp.UnpaidBillofExchange)) oBP.UnpaidBillofExchange = bp.UnpaidBillofExchange;
                if (!String.IsNullOrEmpty(bp.VatGroup)) oBP.VatGroup = bp.VatGroup;
                if (!String.IsNullOrEmpty(bp.VatGroupLatinAmerica)) oBP.VatGroupLatinAmerica = bp.VatGroupLatinAmerica;
                if (!String.IsNullOrEmpty(bp.VatIDNum)) oBP.VatIDNum = bp.VatIDNum;
                oBP.VatLiable = (SAPbobsCOM.BoVatStatus)bp.VatLiable;
                if (!String.IsNullOrEmpty(bp.VATRegistrationNumber)) oBP.VATRegistrationNumber = bp.VATRegistrationNumber;
                if (!String.IsNullOrEmpty(bp.VerificationNumber)) oBP.VerificationNumber = bp.VerificationNumber;
                if (!String.IsNullOrEmpty(bp.Website)) oBP.Website = bp.Website;
                if (bp.WithholdingTaxCertified) oBP.WithholdingTaxCertified = SAPbobsCOM.BoYesNoEnum.tYES;
                if (bp.WithholdingTaxDeductionGroup != 0) oBP.WithholdingTaxDeductionGroup = bp.WithholdingTaxDeductionGroup;
                if (!String.IsNullOrEmpty(bp.WTCode)) oBP.WTCode = bp.WTCode;
                if (!String.IsNullOrEmpty(bp.ZipCode)) oBP.ZipCode = bp.ZipCode;

                foreach (PropertyInfo info in bp.UserFields.GetType().GetProperties())
                {
                    if (bp.UserFields.GetType().GetProperty(info.Name).GetValue(bp.UserFields) != null)
                    {
                        oBP.UserFields.Fields.Item(info.Name).Value = bp.UserFields.GetType().GetProperty(info.Name).GetValue(bp.UserFields);
                    }
                }

                #endregion
                #region Address
                if (bp.Addresses != null)
                {
                    for (int i = 0; i < bp.Addresses.Count; i++)
                    {
                        if (i > 0) oBP.Addresses.Add();
                        oBP.Addresses.SetCurrentLine(i);
                        oBP.Addresses.AddressName = bp.Addresses[i].AddressName;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].AddressName2)) oBP.Addresses.AddressName2 = bp.Addresses[i].AddressName2;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].AddressName3)) oBP.Addresses.AddressName3 = bp.Addresses[i].AddressName3;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].Block)) oBP.Addresses.Block = bp.Addresses[i].Block;
                        oBP.Addresses.AddressType = (SAPbobsCOM.BoAddressType)bp.Addresses[i].AddressType;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].BuildingFloorRoom)) oBP.Addresses.BuildingFloorRoom = bp.Addresses[i].BuildingFloorRoom;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].City)) oBP.Addresses.City = bp.Addresses[i].City;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].Country)) oBP.Addresses.Country = bp.Addresses[i].Country;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].County)) oBP.Addresses.County = bp.Addresses[i].County;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].FederalTaxID)) oBP.Addresses.FederalTaxID = bp.Addresses[i].FederalTaxID;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].GlobalLocationNumber)) oBP.Addresses.GlobalLocationNumber = bp.Addresses[i].GlobalLocationNumber;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].State)) oBP.Addresses.State = bp.Addresses[i].State;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].Street)) oBP.Addresses.Street = bp.Addresses[i].Street;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].StreetNo)) oBP.Addresses.StreetNo = bp.Addresses[i].StreetNo;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].TaxCode)) oBP.Addresses.TaxCode = bp.Addresses[i].TaxCode;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].TypeOfAddress)) oBP.Addresses.TypeOfAddress = bp.Addresses[i].TypeOfAddress;
                        if (!String.IsNullOrEmpty(bp.Addresses[i].ZipCode)) oBP.Addresses.ZipCode = bp.Addresses[i].ZipCode;

                        foreach (PropertyInfo info in bp.Addresses[i].UserFields.GetType().GetProperties())
                        {
                            if (bp.Addresses[i].UserFields.GetType().GetProperty(info.Name).GetValue(bp.Addresses[i].UserFields) != null)
                            {
                                oBP.Addresses.UserFields.Fields.Item(info.Name).Value = bp.Addresses[i].UserFields.GetType().GetProperty(info.Name).GetValue(bp.Addresses[i].UserFields);
                            }
                        }
                    }
                }
                #endregion
                #region Contact Person
                if (bp.Contacts != null)
                {
                    for (int i = 0; i < bp.Contacts.Count; i++)
                    {
                        if (i > 0) oBP.ContactEmployees.Add();
                        oBP.ContactEmployees.SetCurrentLine(i);
                        oBP.ContactEmployees.Name = bp.Contacts[i].Name;
                        if (bp.Contacts[i].Active) oBP.ContactEmployees.Active = SAPbobsCOM.BoYesNoEnum.tYES;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].Address)) oBP.ContactEmployees.Address = bp.Contacts[i].Address;
                        if (bp.Contacts[i].DateOfBirth != DateTime.MinValue) oBP.ContactEmployees.DateOfBirth = bp.Contacts[i].DateOfBirth;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].CityOfBirth)) oBP.ContactEmployees.CityOfBirth = bp.Contacts[i].CityOfBirth;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].E_Mail)) oBP.ContactEmployees.E_Mail = bp.Contacts[i].E_Mail;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].Fax)) oBP.ContactEmployees.Fax = bp.Contacts[i].Fax;
                        oBP.ContactEmployees.Gender = (SAPbobsCOM.BoGenderTypes)bp.Contacts[i].Gender;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].FirstName)) oBP.ContactEmployees.FirstName = bp.Contacts[i].FirstName;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].LastName)) oBP.ContactEmployees.LastName = bp.Contacts[i].LastName;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].MiddleName)) oBP.ContactEmployees.MiddleName = bp.Contacts[i].MiddleName;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].MobilePhone)) oBP.ContactEmployees.MobilePhone = bp.Contacts[i].MobilePhone;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].Pager)) oBP.ContactEmployees.Pager = bp.Contacts[i].Pager;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].Password)) oBP.ContactEmployees.Password = bp.Contacts[i].Password;
                        if (String.IsNullOrEmpty(bp.Contacts[i].Phone1)) oBP.ContactEmployees.Phone1 = bp.Contacts[i].Phone1;
                        if (String.IsNullOrEmpty(bp.Contacts[i].Phone2)) oBP.ContactEmployees.Phone2 = bp.Contacts[i].Phone2;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].PlaceOfBirth)) oBP.ContactEmployees.PlaceOfBirth = bp.Contacts[i].PlaceOfBirth;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].Position)) oBP.ContactEmployees.Position = bp.Contacts[i].Position;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].Profession)) oBP.ContactEmployees.Profession = bp.Contacts[i].Profession;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].Remarks1)) oBP.ContactEmployees.Remarks1 = bp.Contacts[i].Remarks1;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].Remarks2)) oBP.ContactEmployees.Remarks2 = bp.Contacts[i].Remarks2;
                        if (!String.IsNullOrEmpty(bp.Contacts[i].Title)) oBP.ContactEmployees.Title = bp.Contacts[i].Title;
                        foreach (PropertyInfo info in bp.Contacts[i].UserFields.GetType().GetProperties())
                        {
                            if (bp.Contacts[i].UserFields.GetType().GetProperty(info.Name).GetValue(bp.Contacts[i].UserFields) != null)
                            {
                                oBP.ContactEmployees.UserFields.Fields.Item(info.Name).Value = bp.Contacts[i].UserFields.GetType().GetProperty(info.Name).GetValue(bp.Contacts[i].UserFields);
                            }
                        }
                    }
                }
                #endregion

                if (oBP.Add() != 0)
                {
                    errMsg = oCom.GetLastErrorDescription();
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oBP = null;
                    return false;
                }
                oCom.GetNewObjectCode(out key);
                if (key == "")
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oBP = null;
                    return false;
                }
                SAPbobsCOM.Items getDoc = (SAPbobsCOM.Items)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);

                if (!getDoc.GetByKey(key))
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oBP = null;
                    return false;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBP);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(getDoc);
                oBP = null;
                getDoc = null;

                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                errMsg = ex.Message;
                return false;
            }
            return true;
        }

        public bool CreateJournalEntry(Models.SAPModels.JournalEntry je, ref string key)
        {
            //if (!generateLog(je, "30", je.LogUserID)) return false;
            try
            {
                oCom.StartTransaction();
                SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                #region Header
                if (je.Series != 0) oJE.Series = je.Series;
                if (je.AutoVAT) oJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES; else oJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tNO;
                if (je.DeferredTax) oJE.DeferredTax = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(je.DocumentType)) oJE.DocumentType = je.DocumentType;
                if (je.DueDate != DateTime.MinValue) oJE.DueDate = je.DueDate;
                if (!String.IsNullOrEmpty(je.Memo)) oJE.Memo = je.Memo;
                if (!String.IsNullOrEmpty(je.ProjectCode)) oJE.ProjectCode = je.ProjectCode;
                if (!String.IsNullOrEmpty(je.Reference)) oJE.Reference = je.Reference;
                if (!String.IsNullOrEmpty(je.Reference2)) oJE.Reference2 = je.Reference2;
                if (!String.IsNullOrEmpty(je.Reference3)) oJE.Reference3 = je.Reference3;
                if (je.ReferenceDate != DateTime.MinValue) oJE.ReferenceDate = je.ReferenceDate;
                if (je.TaxDate != DateTime.MinValue) oJE.TaxDate = je.TaxDate;
                if (!String.IsNullOrEmpty(je.TransactionCode)) oJE.TransactionCode = je.TransactionCode;

                foreach (PropertyInfo info in je.UserFields.GetType().GetProperties())
                {
                    if (je.UserFields.GetType().GetProperty(info.Name).GetValue(je.UserFields) != null)
                    {
                        oJE.UserFields.Fields.Item(info.Name).Value = je.UserFields.GetType().GetProperty(info.Name).GetValue(je.UserFields);
                    }
                }

                #endregion

                #region Lines
                if (je.Lines != null)
                {
                    for (int i = 0; i < je.Lines.Count; i++)
                    {
                        if (i > 0) oJE.Lines.Add();
                        oJE.Lines.SetCurrentLine(i);

                        if (!String.IsNullOrEmpty(je.Lines[i].AccountCode)) oJE.Lines.AccountCode = je.Lines[i].AccountCode;
                        if (!String.IsNullOrEmpty(je.Lines[i].AdditionalReference)) oJE.Lines.AdditionalReference = je.Lines[i].AdditionalReference;
                        if (je.Lines[i].BaseSum != 0) oJE.Lines.BaseSum = je.Lines[i].BaseSum;
                        if (je.Lines[i].BlockReason != 0) oJE.Lines.BlockReason = je.Lines[i].BlockReason;
                        if (je.Lines[i].BPLID != 0) oJE.Lines.BPLID = je.Lines[i].BPLID;
                        if (!String.IsNullOrEmpty(je.Lines[i].ContraAccount)) oJE.Lines.ContraAccount = je.Lines[i].ContraAccount;
                        if (!String.IsNullOrEmpty(je.Lines[i].ControlAccount)) oJE.Lines.ControlAccount = je.Lines[i].ControlAccount;
                        if (!String.IsNullOrEmpty(je.Lines[i].CostingCode)) oJE.Lines.CostingCode = je.Lines[i].CostingCode;
                        if (!String.IsNullOrEmpty(je.Lines[i].CostingCode2)) oJE.Lines.CostingCode2 = je.Lines[i].CostingCode2;
                        if (!String.IsNullOrEmpty(je.Lines[i].CostingCode3)) oJE.Lines.CostingCode3 = je.Lines[i].CostingCode3;
                        if (!String.IsNullOrEmpty(je.Lines[i].CostingCode4)) oJE.Lines.CostingCode4 = je.Lines[i].CostingCode4;
                        if (!String.IsNullOrEmpty(je.Lines[i].CostingCode5)) oJE.Lines.CostingCode5 = je.Lines[i].CostingCode5;
                        if (je.Lines[i].Credit != 0) oJE.Lines.Credit = je.Lines[i].Credit;
                        if (je.Lines[i].Debit != 0) oJE.Lines.Debit = je.Lines[i].Debit;
                        if (je.Lines[i].DueDate != DateTime.MinValue) oJE.Lines.DueDate = je.Lines[i].DueDate;
                        if (je.Lines[i].FCCredit != 0) oJE.Lines.FCCredit = je.Lines[i].FCCredit;
                        if (je.Lines[i].FCDebit != 0) oJE.Lines.FCDebit = je.Lines[i].FCDebit;
                        if (!String.IsNullOrEmpty(je.Lines[i].FCCurrency)) oJE.Lines.FCCurrency = je.Lines[i].FCCurrency;
                        if (!String.IsNullOrEmpty(je.Lines[i].LineMemo)) oJE.Lines.LineMemo = je.Lines[i].LineMemo;
                        if (!String.IsNullOrEmpty(je.Lines[i].ProjectCode)) oJE.Lines.ProjectCode = je.Lines[i].ProjectCode;
                        if (!String.IsNullOrEmpty(je.Lines[i].Reference1)) oJE.Lines.Reference1 = je.Lines[i].Reference1;
                        if (!String.IsNullOrEmpty(je.Lines[i].Reference2)) oJE.Lines.Reference2 = je.Lines[i].Reference2;
                        if (je.Lines[i].ReferenceDate1 != DateTime.MinValue) oJE.Lines.ReferenceDate1 = je.Lines[i].ReferenceDate1;
                        if (je.Lines[i].ReferenceDate2 != DateTime.MinValue) oJE.Lines.ReferenceDate2 = je.Lines[i].ReferenceDate2;
                        if (!String.IsNullOrEmpty(je.Lines[i].ShortName)) oJE.Lines.ShortName = je.Lines[i].ShortName;
                        if (!String.IsNullOrEmpty(je.Lines[i].TaxCode)) oJE.Lines.TaxCode = je.Lines[i].TaxCode;
                        if (!String.IsNullOrEmpty(je.Lines[i].TaxGroup)) oJE.Lines.TaxGroup = je.Lines[i].TaxGroup;
                        if (je.Lines[i].VatAmount != 0) oJE.Lines.VatAmount = je.Lines[i].VatAmount;

                        foreach (PropertyInfo info in je.Lines[i].UserFields.GetType().GetProperties())
                        {
                            if (je.Lines[i].UserFields.GetType().GetProperty(info.Name).GetValue(je.Lines[i].UserFields) != null)
                            {
                                oJE.Lines.UserFields.Fields.Item(info.Name).Value = je.Lines[i].UserFields.GetType().GetProperty(info.Name).GetValue(je.Lines[i].UserFields);
                            }
                        }
                    }
                }
                #endregion

                if (oJE.Add() != 0)
                {
                    errMsg = oCom.GetLastErrorDescription();
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oJE = null;
                    return false;
                }
                oCom.GetNewObjectCode(out key);
                if (key == "")
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oJE = null;
                    return false;
                }
                SAPbobsCOM.JournalEntries getDoc = (SAPbobsCOM.JournalEntries)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                if (!getDoc.GetByKey(int.Parse(key)))
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oJE = null;
                    return false;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oJE);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(getDoc);
                oJE = null;
                getDoc = null;

                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                errMsg = ex.Message;
                return false;
            }

            return true;
        }

        public bool CreateJournalVoucher(Models.SAPModels.JournalEntry je, ref string key)
        {
            //if (!generateLog(je, "28", je.LogUserID)) return false;
            oCom.StartTransaction();
            try
            {
                SAPbobsCOM.JournalVouchers oJV = (SAPbobsCOM.JournalVouchers)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers);

                SAPbobsCOM.JournalEntries oJE = oJV.JournalEntries;
                #region Header
                if (je.Series != 0) oJE.Series = je.Series;
                if (je.AutoVAT) oJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tYES; else oJE.AutoVAT = SAPbobsCOM.BoYesNoEnum.tNO;
                if (je.DeferredTax) oJE.DeferredTax = SAPbobsCOM.BoYesNoEnum.tYES;
                if (!String.IsNullOrEmpty(je.DocumentType)) oJE.DocumentType = je.DocumentType;
                if (je.DueDate != DateTime.MinValue) oJE.DueDate = je.DueDate;
                if (String.IsNullOrEmpty(je.Memo)) oJE.Memo = je.Memo;
                if (!String.IsNullOrEmpty(je.ProjectCode)) oJE.ProjectCode = je.ProjectCode;
                if (!String.IsNullOrEmpty(je.Reference)) oJE.Reference = je.Reference;
                if (!String.IsNullOrEmpty(je.Reference2)) oJE.Reference2 = je.Reference2;
                if (!String.IsNullOrEmpty(je.Reference3)) oJE.Reference3 = je.Reference3;
                if (je.ReferenceDate != DateTime.MinValue) oJE.ReferenceDate = je.ReferenceDate;
                if (je.TaxDate != DateTime.MinValue) oJE.TaxDate = je.TaxDate;
                if (!String.IsNullOrEmpty(je.TransactionCode)) oJE.TransactionCode = je.TransactionCode;

                foreach (PropertyInfo info in je.UserFields.GetType().GetProperties())
                {
                    if (je.UserFields.GetType().GetProperty(info.Name).GetValue(je.UserFields) != null)
                    {
                        oJE.UserFields.Fields.Item(info.Name).Value = je.UserFields.GetType().GetProperty(info.Name).GetValue(je.UserFields);
                    }
                }

                #endregion

                #region Lines
                if (je.Lines != null)
                {
                    for (int i = 0; i < je.Lines.Count; i++)
                    {
                        if (i > 0) oJE.Lines.Add();
                        oJE.Lines.SetCurrentLine(i);
                        if (!String.IsNullOrEmpty(je.Lines[i].AccountCode)) oJE.Lines.AccountCode = je.Lines[i].AccountCode;
                        if (!String.IsNullOrEmpty(je.Lines[i].AdditionalReference)) oJE.Lines.AdditionalReference = je.Lines[i].AdditionalReference;
                        if (je.Lines[i].BaseSum != 0) oJE.Lines.BaseSum = je.Lines[i].BaseSum;
                        if (je.Lines[i].BlockReason != 0) oJE.Lines.BlockReason = je.Lines[i].BlockReason;
                        if (je.Lines[i].BPLID != 0) oJE.Lines.BPLID = je.Lines[i].BPLID;
                        if (!String.IsNullOrEmpty(je.Lines[i].ContraAccount)) oJE.Lines.ContraAccount = je.Lines[i].ContraAccount;
                        if (!String.IsNullOrEmpty(je.Lines[i].ControlAccount)) oJE.Lines.ControlAccount = je.Lines[i].ControlAccount;
                        if (!String.IsNullOrEmpty(je.Lines[i].CostingCode)) oJE.Lines.CostingCode = je.Lines[i].CostingCode;
                        if (!String.IsNullOrEmpty(je.Lines[i].CostingCode2)) oJE.Lines.CostingCode2 = je.Lines[i].CostingCode2;
                        if (!String.IsNullOrEmpty(je.Lines[i].CostingCode3)) oJE.Lines.CostingCode3 = je.Lines[i].CostingCode3;
                        if (!String.IsNullOrEmpty(je.Lines[i].CostingCode4)) oJE.Lines.CostingCode4 = je.Lines[i].CostingCode4;
                        if (!String.IsNullOrEmpty(je.Lines[i].CostingCode5)) oJE.Lines.CostingCode5 = je.Lines[i].CostingCode5;
                        if (je.Lines[i].Credit != 0) oJE.Lines.Credit = je.Lines[i].Credit;
                        if (je.Lines[i].Debit != 0) oJE.Lines.Debit = je.Lines[i].Debit;
                        if (je.Lines[i].DueDate != DateTime.MinValue) oJE.Lines.DueDate = je.Lines[i].DueDate;
                        if (je.Lines[i].FCCredit != 0) oJE.Lines.FCCredit = je.Lines[i].FCCredit;
                        if (je.Lines[i].FCDebit != 0) oJE.Lines.FCDebit = je.Lines[i].FCDebit;
                        if (!String.IsNullOrEmpty(je.Lines[i].FCCurrency)) oJE.Lines.FCCurrency = je.Lines[i].FCCurrency;
                        if (!String.IsNullOrEmpty(je.Lines[i].LineMemo)) oJE.Lines.LineMemo = je.Lines[i].LineMemo;
                        if (!String.IsNullOrEmpty(je.Lines[i].ProjectCode)) oJE.Lines.ProjectCode = je.Lines[i].ProjectCode;
                        if (!String.IsNullOrEmpty(je.Lines[i].Reference1)) oJE.Lines.Reference1 = je.Lines[i].Reference1;
                        if (!String.IsNullOrEmpty(je.Lines[i].Reference2)) oJE.Lines.Reference2 = je.Lines[i].Reference2;
                        if (je.Lines[i].ReferenceDate1 != DateTime.MinValue) oJE.Lines.ReferenceDate1 = je.Lines[i].ReferenceDate1;
                        if (je.Lines[i].ReferenceDate2 != DateTime.MinValue) oJE.Lines.ReferenceDate2 = je.Lines[i].ReferenceDate2;
                        if (!String.IsNullOrEmpty(je.Lines[i].ShortName)) oJE.Lines.ShortName = je.Lines[i].ShortName;
                        if (!String.IsNullOrEmpty(je.Lines[i].TaxCode)) oJE.Lines.TaxCode = je.Lines[i].TaxCode;
                        if (!String.IsNullOrEmpty(je.Lines[i].TaxGroup)) oJE.Lines.TaxGroup = je.Lines[i].TaxGroup;
                        if (je.Lines[i].VatAmount != 0) oJE.Lines.VatAmount = je.Lines[i].VatAmount;

                        foreach (PropertyInfo info in je.Lines[i].UserFields.GetType().GetProperties())
                        {
                            if (je.Lines[i].UserFields.GetType().GetProperty(info.Name).GetValue(je.Lines[i].UserFields) != null)
                            {
                                oJE.Lines.UserFields.Fields.Item(info.Name).Value = je.Lines[i].UserFields.GetType().GetProperty(info.Name).GetValue(je.Lines[i].UserFields);
                            }
                        }
                    }
                }
                #endregion

                if (oJV.Add() != 0)
                {
                    errMsg = oCom.GetLastErrorDescription();
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oJE = null;
                    oJV = null;
                    return false;
                }
                oCom.GetNewObjectCode(out key);
                if (key == "")
                {
                    errMsg = "Unknown Error! Please try again!";
                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    oJE = null;
                    oJV = null;
                    return false;
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oJE);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oJV);
                oJE = null;
                oJV = null;

                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {
                if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                errMsg = ex.Message;
                return false;
            }

            return true;
        }
        public DataSet getDocumentDataset(string docEntry, string ObjectType)
        {
            DataSet ds = new DataSet();
            string headerTable = "", detailtable = "";
            switch (ObjectType)
            {
                case "13":
                    headerTable = "OINV";
                    detailtable = "INV1";
                    break;
                case "17":
                    headerTable = "ORDR";
                    detailtable = "RDR1";
                    break;
            }
            using (SqlConnection conn = new SqlConnection(getConnectionString()))
            {
                using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
                {
                    DataTable dt = new DataTable(headerTable);
                    DataTable dt1 = new DataTable(detailtable);
                    dat.SelectCommand.CommandText = "SELECT * FROM " + headerTable + " WHERE DOCENTRY = @1";
                    dat.SelectCommand.Parameters.Clear();
                    dat.SelectCommand.Parameters.AddWithValue("@1", docEntry);
                    dat.Fill(dt);

                    dat.SelectCommand.CommandText = "SELECT * FROM " + detailtable + " WHERE DOCENTRY = @1";
                    dat.SelectCommand.Parameters.Clear();
                    dat.SelectCommand.Parameters.AddWithValue("@1", docEntry);
                    dat.Fill(dt1);
                    ds.Tables.Add(dt);
                    ds.Tables.Add(dt1);
                }
            }
            return ds;

        }

        public DataSet getBPList(string cardType, string cardCode)
        {
            if (String.IsNullOrEmpty(cardType)) cardType = "";
            if (String.IsNullOrEmpty(cardCode)) cardCode = "";
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(getConnectionString()))
            {
                using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
                {
                    DataTable dt = new DataTable("OCRD");

                    string selectcmd = "SELECT * FROM OCRD WHERE 0 = 0 ";
                    if (cardType != "") selectcmd += " AND CardType = @CardType ";
                    if (cardCode != "")
                    {
                        selectcmd += " AND CardCode = @CardCode ";
                    }

                    dat.SelectCommand.CommandText = selectcmd;
                    dat.SelectCommand.Parameters.Clear();
                    if (selectcmd.Contains("@CardType")) dat.SelectCommand.Parameters.AddWithValue("@CardType", cardType);
                    if (selectcmd.Contains("@CardCode")) dat.SelectCommand.Parameters.AddWithValue("@CardCode", cardCode);
                    dat.Fill(dt);

                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }
        public DataSet getItemList(string whsCode, string itemCode)
        {
            if (String.IsNullOrEmpty(whsCode)) whsCode = "";
            if (String.IsNullOrEmpty(itemCode)) itemCode = "";
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(getConnectionString()))
            {
                using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
                {
                    DataTable dt = new DataTable("OITM");
                    DataTable dt1 = new DataTable("OITW");
                    string selectcmd = "SELECT * FROM OITM WHERE FrozenFor = 'N' ";
                    if (itemCode != "") selectcmd += " AND ItemCode = @ItemCode ";

                    dat.SelectCommand.CommandText = selectcmd;
                    dat.SelectCommand.Parameters.Clear();
                    if (selectcmd.Contains("@ItemCode")) dat.SelectCommand.Parameters.AddWithValue("@ItemCode", itemCode);
                    dat.Fill(dt);



                    selectcmd = "SELECT * FROM OITW WHERE 0 = 0 ";
                    if (itemCode != "") selectcmd += " AND ItemCode = @ItemCode ";
                    if (whsCode != "") selectcmd += " AND WhsCode = @WhsCode ";
                    dat.SelectCommand.CommandText = selectcmd;
                    dat.SelectCommand.Parameters.Clear();
                    if (selectcmd.Contains("@ItemCode")) dat.SelectCommand.Parameters.AddWithValue("@ItemCode", itemCode);
                    if (selectcmd.Contains("@WhsCode")) dat.SelectCommand.Parameters.AddWithValue("@WhsCode", whsCode);
                    dat.Fill(dt1);

                    ds.Tables.Add(dt);
                    ds.Tables.Add(dt1);
                }
            }
            return ds;
        }
        public DataSet getItemListWithPrice(string cardCode, string itemCode)
        {
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(getConnectionString()))
            {
                using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
                {
                    DataTable dt = new DataTable("OITM");
                    dat.SelectCommand.CommandText = "select T0.ItemCode, T0.ItemName , T1.Price " +
                        "from OITM T0 " +
                        "LEFT OUTER JOIN ITM1 T1 INNER JOIN OCRD T2 " +
                        "ON T2.ListNum = T1.PriceList AND T2.CardCode = @CardCode ON T1.ItemCode = T0.ItemCode " +
                        "WHERE T0.ItemCode = CASE WHEN ISNULL(@ItemCode,'') = '' THEN T0.ItemCode ELSE @ItemCode END ";
                    dat.SelectCommand.Parameters.Clear();
                    dat.SelectCommand.Parameters.AddWithValue("@ItemCode", itemCode);
                    dat.SelectCommand.Parameters.AddWithValue("@CardCode", cardCode);
                    dat.Fill(dt);
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }
        public DataSet getStatement(string docDate, string cardCode)
        {

            string selectcmd = "DECLARE @AgeBy NVARCHAR(10) " +
                "SET @AgeBy = 'DOC' " +
                "SELECT @CardCodeFrom = ISNULL(@CardCodeFrom,''), @CardCodeTo = ISNULL(@CardCodeTo,'') " +
                "DECLARE @StatementMthStart SMALLDATETIME " +
                "SET @StatementMthStart = DATEADD(mm, DATEDIFF(mm, 0, @StatementDate), 0) " +
                "DECLARE @Maincurncy VARCHAR(3) " +
                "declare @Maincurrname as nvarchar(20) " +
                "DECLARE @CompnyName nvarchar(100) " +
                "DECLARE @CompnyAddr nvarchar(254) " +
                "SELECT @Maincurncy = OADM.Maincurncy, @Maincurrname = OCRN.currname, @CompnyName = OADM.CompnyName, @CompnyAddr = OADM.CompnyAddr FROM OADM inner join OCRN on OCRN.currcode=OADM.MainCurncy " +
                "IF OBJECT_ID('tempdb..#TempStatement') IS NOT NULL  " +
                "    DROP TABLE #TempStatement " +
                "SELECT T1.TransType,  " +
                "CASE T1.TransType WHEN '13' THEN 'IN' WHEN '14' THEN 'CN' WHEN '24' THEN 'RC' WHEN '46' THEN 'PS' WHEN '30' THEN 'JE'  " +
                "WHEN '203' THEN 'DP' WHEN '204' THEN 'DP' WHEN '-5' THEN 'RU' ELSE '' END AS [DocType], " +
                "T0.TransID, T1.Line_ID AS [TransRowID], T1.CreatedBy AS [DocEntry], T0.DocSeries,  " +
                "T0.BaseRef AS [DocNum], CASE WHEN ISNULL(T1.Ref1,'') = '' THEN CAST(T0.Number AS NVARCHAR(100)) ELSE T1.Ref1 END AS [Ref1], T1.Ref2, T1.RefDate, T1.DueDate,  " +
                "CASE WHEN @AgeBy = 'DUE' THEN T1.DueDate ELSE T1.RefDate END AS [AgeDate],  " +
                "T1.ShortName AS [CardCode], T1.LineMemo, CASE WHEN T1.FCCurrency IS NULL OR T1.FCCurrency = '' THEN @Maincurncy ELSE T1.FCCurrency END AS [DocCurr], " +
                "MAX(T1.Debit - T1.Credit) AS [BaseLCAmount], " +
                "MAX(T1.FCDebit - T1.FCCredit) AS [BaseFCAmount], " +
                "MAX(T1.SysDeb - T1.SysCred) AS [BaseSyAmount], " +
                "-(MAX(T1.BalDueCred) + SUM(TR1.ReconSum)) AS [LCAmount],  " +
                "-(MAX(T1.BalFcCred) + SUM(TR1.ReconSumFC)) AS [FCAmount],  " +
                "-(MAX(T1.BalScCred) + SUM(TR1.ReconSumSC)) AS [SyAmount] " +
                "INTO #TempStatement " +
                "FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.TransID = T1.TransID " +
                "INNER JOIN ITR1 TR1 ON T1.TransID = TR1.TransID AND T1.Line_ID = TR1.TransRowID AND TR1.IsCredit = 'C'   " +
                "INNER JOIN OITR TR0 ON TR1.ReconNum = TR0.ReconNum AND TR0.ReconDate > @StatementDate   " +
                "INNER JOIN OCRD T4 ON T4.CardCode = T1.ShortName and T4.CardType = 'C' " +
                "WHERE T0.RefDate <= @StatementDate   " +
                "AND (T4.CardCode >= @CardCodeFrom OR @CardCodeFrom = '')   " +
                "AND (T4.CardCode <= @CardCodeTo OR @CardCodeTo = '')    " +
                "GROUP BY T1.TransType, T0.TransID, T1.Line_ID, T1.CreatedBy, T0.DocSeries, T0.BaseRef, T1.Ref1, T0.Number, T1.Ref2, T1.DueDate, T1.RefDate, T1.ShortName, T1.LineMemo, T1.FCCurrency " +
                "HAVING MAX(T1.BalFcCred) <> -SUM(TR1.ReconSumFC)  OR  MAX(T1.BalDueCred) <> -SUM(TR1.ReconSum)   " +
                "CREATE INDEX [IDX_CardCode] ON #TempStatement (CardCode) " +
                "INSERT INTO #TempStatement " +
                "SELECT T1.TransType,  " +
                "CASE T1.TransType WHEN '13' THEN 'IN' WHEN '14' THEN 'CN' WHEN '24' THEN 'RC' WHEN '46' THEN 'PS' WHEN '30' THEN 'JE'  " +
                "WHEN '203' THEN 'DP' WHEN '204' THEN 'DP' WHEN '-5' THEN 'RU' ELSE '' END AS [DocType], " +
                "T0.TransID, T1.Line_ID AS [TransRowID], T1.CreatedBy AS [DocEntry], T0.DocSeries,  " +
                "T0.BaseRef AS [DocNum], CASE WHEN ISNULL(T1.Ref1,'') = '' THEN CAST(T0.Number AS NVARCHAR(100)) ELSE T1.Ref1 END AS [Ref1], T1.Ref2, T1.RefDate, T1.DueDate,  " +
                "CASE WHEN @AgeBy = 'DUE' THEN T1.DueDate ELSE T1.RefDate END AS [AgeDate], " +
                "T1.ShortName AS [CardCode], T1.LineMemo, CASE WHEN T1.FCCurrency IS NULL OR T1.FCCurrency = '' THEN @Maincurncy ELSE T1.FCCurrency END AS [DocCurr], " +
                "MAX(T1.Debit - T1.Credit) AS [BaseLCAmount], " +
                "MAX(T1.FCDebit - T1.FCCredit) AS [BaseFCAmount], " +
                "MAX(T1.SysDeb - T1.SysCred) AS [BaseSyAmount], " +
                "(MAX(T1.BalDueDeb) + SUM(TR1.ReconSum)) AS [LCAmount],  " +
                "(MAX(T1.BalFcDeb) + SUM(TR1.ReconSumFC)) AS [FCAmount],  " +
                "(MAX(T1.BalScDeb) + SUM(TR1.ReconSumSC)) AS [SyAmount] " +
                "FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.TransID = T1.TransID " +
                "INNER JOIN ITR1 TR1 ON T1.TransID = TR1.TransID AND T1.Line_ID = TR1.TransRowID AND TR1.IsCredit = 'D'   " +
                "INNER JOIN OITR TR0 ON TR1.ReconNum = TR0.ReconNum AND TR0.ReconDate > @StatementDate   " +
                "INNER JOIN OCRD T4 ON T4.CardCode = T1.ShortName and T4.CardType = 'C' " +
                "WHERE T0.RefDate <= @StatementDate   " +
                "AND (T4.CardCode >= @CardCodeFrom OR @CardCodeFrom = '')   " +
                "AND (T4.CardCode <= @CardCodeTo OR @CardCodeTo = '')   " +
                "GROUP BY T1.TransType, T0.TransID, T1.Line_ID, T1.CreatedBy, T0.DocSeries, T0.BaseRef, T1.Ref1, T0.Number, T1.Ref2, T1.DueDate, T1.RefDate, T1.ShortName, T1.LineMemo, T1.FCCurrency " +
                "HAVING MAX(T1.BalFcDeb) <> -SUM(TR1.ReconSumFC)  OR  MAX(T1.BalDueDeb) <> -SUM(TR1.ReconSum)   " +
                "INSERT INTO #TempStatement " +
                "SELECT T1.TransType,  " +
                "CASE T1.TransType WHEN '13' THEN 'IN' WHEN '14' THEN 'CN' WHEN '24' THEN 'RC' WHEN '46' THEN 'PS' WHEN '30' THEN 'JE'  " +
                "WHEN '203' THEN 'DP' WHEN '204' THEN 'DP' WHEN '-5' THEN 'RU' ELSE '' END AS [DocType], " +
                "T0.TransID, T1.Line_ID AS [TransRowID], T1.CreatedBy AS [DocEntry], T0.DocSeries,  " +
                "T0.BaseRef AS [DocNum], CASE WHEN ISNULL(T1.Ref1,'') = '' THEN CAST(T0.Number AS NVARCHAR(100)) ELSE T1.Ref1 END AS [Ref1], T1.Ref2, T1.RefDate, T1.DueDate,  " +
                "CASE WHEN @AgeBy = 'DUE' THEN T1.DueDate ELSE T1.RefDate END AS [AgeDate], " +
                "T1.ShortName AS [CardCode], T1.LineMemo, CASE WHEN T1.FCCurrency IS NULL OR T1.FCCurrency = '' THEN @Maincurncy ELSE T1.FCCurrency END AS [DocCurr], " +
                "MAX(T1.Debit - T1.Credit) AS [BaseLCAmount], " +
                "MAX(T1.FCDebit - T1.FCCredit) AS [BaseFCAmount], " +
                "MAX(T1.SysDeb - T1.SysCred) AS [BaseSyAmount], " +
                "MAX(T1.BalDueDeb) - MAX(T1.BalDueCred) AS [LCAmount],  " +
                "MAX(T1.BalFcDeb) - MAX(T1.BalFcCred) AS [FCAmount],  " +
                "MAX(T1.BalScDeb) - MAX(T1.BalScCred) AS [SyAmount] " +
                "FROM OJDT T0 INNER JOIN JDT1 T1 ON T0.TransID = T1.TransID " +
                "INNER JOIN OCRD T4 ON T4.CardCode = T1.ShortName and T4.CardType = 'C' " +
                "WHERE T0.RefDate <= @StatementDate   " +
                "AND (T4.CardCode >= @CardCodeFrom OR @CardCodeFrom = '')   " +
                "AND (T4.CardCode <= @CardCodeTo OR @CardCodeTo = '')   " +
                "AND NOT EXISTS ( " +
                "    SELECT 1 FROM OITR R0 INNER JOIN ITR1 R1 ON R0.ReconNum = R1.ReconNum " +
                "    WHERE R1.TransID = T1.TransID AND R1.TransRowID = T1.Line_ID AND R0.ReconDate > @StatementDate " +
                ")   " +
                "GROUP BY T1.TransType, T0.TransID, T1.Line_ID, T1.CreatedBy, T0.DocSeries, T0.BaseRef, T1.Ref1, T0.Number, T1.Ref2, T1.DueDate, T1.RefDate, T1.ShortName, T1.LineMemo, T1.FCCurrency " +
                "HAVING MAX(T1.BalFcCred) <> MAX(T1.BalFcDeb)  OR  MAX(T1.BalDueCred) <> MAX(T1.BalDueDeb)   " +
                "UPDATE #TempStatement SET Ref2 = T1.NumAtCard " +
                "FROM #TempStatement T0 INNER JOIN ORIN T1 ON T0.TransType = T1.ObjType collate database_default AND T0.DocEntry = T1.DocEntry " +
                "WHERE T0.TransType = '14' AND T0.Ref2 = '' " +
                "DELETE #TempStatement " +
                "WHERE CardCode IN " +
                "(SELECT CardCode FROM #TempStatement GROUP BY CardCode HAVING SUM(LCAmount) = 0 AND SUM(FCAmount) = 0 AND SUM(SyAmount) = 0) " +
                "AND CardCode NOT IN " +
                "(SELECT CardCode FROM #TempStatement WHERE RefDate >= @StatementMthStart AND RefDate <= @StatementDate GROUP BY CardCode) " +
                "SELECT T0.*,DS.OcrCode, " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) <= 0 THEN T0.LCAmount ELSE 0 END) AS [LCCurrent], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) = 1 THEN T0.LCAmount ELSE 0 END) AS [LCOver30], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) = 2 THEN T0.LCAmount ELSE 0 END) AS [LCOver60], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) = 3 THEN T0.LCAmount ELSE 0 END) AS [LCOver90], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) = 4 THEN T0.LCAmount ELSE 0 END) AS [LCOver120], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) = 5 THEN T0.LCAmount ELSE 0 END) AS [LCOver150], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) > 5 THEN T0.LCAmount ELSE 0 END) AS [LCOver180], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) <= 0 THEN T0.FCAmount ELSE 0 END) AS [FCCurrent], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) = 1 THEN T0.FCAmount ELSE 0 END) AS [FCOver30], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) = 2 THEN T0.FCAmount ELSE 0 END) AS [FCOver60], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) = 3 THEN T0.FCAmount ELSE 0 END) AS [FCOver90], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) = 4 THEN T0.FCAmount ELSE 0 END) AS [FCOver120], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) = 5 THEN T0.FCAmount ELSE 0 END) AS [FCOver150], " +
                "    (CASE WHEN DATEDIFF(m, T0.AgeDate, @StatementDate) > 5 THEN T0.FCAmount ELSE 0 END) AS [FCOver180], " +
                "    T2.CardName, T2.CardFName, T7.GroupName, T2.Phone1, T2.Fax, T2.CreditLine, T2.Currency AS [CardCurr],  " +
                "    CASE WHEN T2.Currency = '##' THEN @Maincurrname ELSE T6.CurrName END AS [CardCurrName], " +
                "    T3.PymntGroup, T2.CntctPrsn, ISNULL(DS.SlpName, T4.SlpName) AS [SlpName], ISNULL(DS.Memo, ISNULL(T4.Memo,'')) AS [SlpMemo], " +
                "    ISNULL(T5.SeriesName,'') AS [SeriesName],  " +
                "    ISNULL(T2.Address,'') AS [Address], ISNULL(T2.Block,'') AS [Address2], ISNULL(T2.City,'') AS [Address3], ISNULL(T2.County,'') AS [Address4],  " +
                "    @Maincurncy AS [Maincurncy], @Maincurrname AS [Maincurrname], @CompnyName AS [CompnyName], @CompnyAddr AS [CompnyAddr], @StatementMthStart AS [StatementMthStart], " +
                "    ISNULL(TS.YTDDocTotal, 0) AS YTDDocTotal, ISNULL(TS.YTDDocTotalFC, 0) AS YTDDocTotalFC,  " +
                "    ISNULL(TS.MTDDocTotal, 0) AS MTDDocTotal, ISNULL(TS.MTDDocTotalFC, 0) AS MTDDocTotalFC " +
                "FROM #TempStatement T0  " +
                "INNER JOIN OCRD T2 ON T0.CardCode = T2.CardCode " +
                "LEFT OUTER JOIN OCTG T3 ON T2.GroupNum = T3.GroupNum " +
                "LEFT OUTER JOIN OSLP T4 ON T2.SlpCode = T4.SlpCode " +
                "LEFT OUTER JOIN NNM1 T5 ON T0.TransType = T5.ObjectCode AND T0.DocSeries = T5.Series " +
                "LEFT OUTER JOIN OCRN T6 ON T2.Currency = T6.CurrCode " +
                "INNER JOIN OCRG T7 ON T2.GroupCode = T7.GroupCode " +
                "LEFT OUTER JOIN " +
                "( " +
                "    SELECT CardCode, SUM(YTDDocTotal) AS YTDDocTotal, SUM(YTDDocTotalFC) AS YTDDocTotalFC " +
                "    , SUM(MTDDocTotal) AS MTDDocTotal, SUM(MTDDocTotalFC) AS MTDDocTotalFC " +
                "    FROM ( " +
                "    SELECT CardCode, T0.DocTotal AS YTDDocTotal, T0.DocTotalFC AS YTDDocTotalFC,  " +
                "    CASE WHEN T0.DocDate >= DATEADD(mm, DATEDIFF(mm, 0, @StatementDate), 0) THEN T0.DocTotal ELSE 0 END AS MTDDocTotal, " +
                "    CASE WHEN T0.DocDate >= DATEADD(mm, DATEDIFF(mm, 0, @StatementDate), 0) THEN T0.DocTotalFC ELSE 0 END AS MTDDocTotalFC " +
                "    FROM  OINV T0  " +
                "    WHERE T0.DocDate >= DATEADD(yy, DATEDIFF(yy, 0, @StatementDate), 0) AND T0.DocDate <= @StatementDate " +
                "    AND (T0.CardCode >= @CardCodeFrom OR @CardCodeFrom = '') AND (T0.CardCode <= @CardCodeTo OR @CardCodeTo = '') " +
                "    UNION ALL " +
                "    SELECT CardCode, -T0.DocTotal, -T0.DocTotalFC,  " +
                "    CASE WHEN T0.DocDate >= DATEADD(mm, DATEDIFF(mm, 0, @StatementDate), 0) THEN -T0.DocTotal ELSE 0 END AS MTDDocTotal, " +
                "    CASE WHEN T0.DocDate >= DATEADD(mm, DATEDIFF(mm, 0, @StatementDate), 0) THEN -T0.DocTotalFC ELSE 0 END AS MTDDocTotalFC " +
                "    FROM  ORIN T0  " +
                "    WHERE T0.DocDate >= DATEADD(yy, DATEDIFF(yy, 0, @StatementDate), 0) AND T0.DocDate <= @StatementDate " +
                "    AND (T0.CardCode >= @CardCodeFrom OR @CardCodeFrom = '') AND (T0.CardCode <= @CardCodeTo OR @CardCodeTo = '') " +
                "    ) RESULT GROUP BY CardCode " +
                ") TS ON T0.CardCode = TS.CardCode  " +
                "LEFT OUTER JOIN " +
                "( " +
                "    SELECT DISTINCT T0.TransID, T1.SlpName, ISNULL(T1.Memo,'') AS [Memo], T2.OcrCode " +
                "    FROM OINV T0 INNER JOIN OSLP T1 ON T0.SlpCode = T1.SlpCode INNER JOIN INV1 T2 ON T2.DocEntry = T0.DocEntry " +
                "    WHERE T0.DocDate <= @StatementDate " +
                "    AND (T0.CardCode >= @CardCodeFrom OR @CardCodeFrom = '') AND (T0.CardCode <= @CardCodeTo OR @CardCodeTo = '') " +
                "    UNION ALL " +
                "    SELECT DISTINCT T0.TransID, T1.SlpName, ISNULL(T1.Memo,'') AS [Memo], T2.OcrCode " +
                "    FROM ORIN T0 INNER JOIN OSLP T1 ON T0.SlpCode = T1.SlpCode INNER JOIN RIN1 T2 ON T2.DocEntry = T0.DocEntry " +
                "    WHERE T0.DocDate <= @StatementDate " +
                "    AND (T0.CardCode >= @CardCodeFrom OR @CardCodeFrom = '') AND (T0.CardCode <= @CardCodeTo OR @CardCodeTo = '') " +
                ") DS ON T0.TransID = DS.TransID " +
                "DROP TABLE #TempStatement ";

            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(getConnectionString()))
            {
                using (SqlDataAdapter dat = new SqlDataAdapter(selectcmd, conn))
                {
                    DataTable dt = new DataTable("Statement");
                    dat.SelectCommand.Parameters.Clear();
                    dat.SelectCommand.Parameters.AddWithValue("@StatementDate", docDate);
                    dat.SelectCommand.Parameters.AddWithValue("@CardCodeFrom", cardCode);
                    dat.SelectCommand.Parameters.AddWithValue("@CardCodeTo", cardCode);
                    dat.Fill(dt);


                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }

        public DataSet getSeries(string docType, string id)
        {
            if (String.IsNullOrEmpty(id)) id = "";
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(getConnectionString()))
            {
                using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
                {
                    DataTable dt = new DataTable("series");
                    string selectcmd = "select * from NNM1 where ObjectCode = @ObjectCode and Locked = 'N' ";
                    if (id != "") selectcmd += " AND Series = @series ";
                    dat.SelectCommand.CommandText = selectcmd;
                    dat.SelectCommand.Parameters.Clear();
                    dat.SelectCommand.Parameters.AddWithValue("@ObjectCode", docType);
                    if (selectcmd.Contains("@series")) dat.SelectCommand.Parameters.AddWithValue("@series", id);
                    dat.Fill(dt);

                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }
        //public bool createAceSO(SAPB1WebApi.Ace.AceDoc doc, ref string docentry)
        //{
        //    if (!generateLog(doc, "17", "")) return false;
        //    try
        //    {
        //        oCom.StartTransaction();
        //        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
        //        #region Header
        //        if (doc.HandWritten)
        //        {
        //            oDoc.DocNum = doc.DocNum;
        //            oDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES;
        //        }
        //        else
        //        {
        //            if (doc.Series > 0) oDoc.Series = doc.Series;
        //        }
        //        oDoc.DocDate = doc.DocDate;
        //        oDoc.CardCode = doc.CardCode;
        //        if (!String.IsNullOrEmpty(doc.ClosingRemarks)) oDoc.ClosingRemarks = doc.ClosingRemarks;
        //        if (!String.IsNullOrEmpty(doc.Comments)) oDoc.Comments = doc.Comments;
        //        if (doc.DiscountPercent != 0) oDoc.DiscountPercent = doc.DiscountPercent;
        //        if (!String.IsNullOrEmpty(doc.DocCurrency)) oDoc.DocCurrency = doc.DocCurrency;
        //        if (doc.DocDueDate != DateTime.MinValue) oDoc.DocDueDate = doc.DocDueDate;
        //        if (doc.DocRate > 0) oDoc.DocRate = doc.DocRate;
        //        if (doc.DocType > 0) oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service; else oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
        //        if (!String.IsNullOrEmpty(doc.JournalMemo)) oDoc.JournalMemo = doc.JournalMemo;
        //        if (!String.IsNullOrEmpty(doc.NumAtCard)) oDoc.NumAtCard = doc.NumAtCard;
        //        if (!String.IsNullOrEmpty(doc.OpeningRemarks)) oDoc.OpeningRemarks = doc.OpeningRemarks;
        //        if (!String.IsNullOrEmpty(doc.Project)) oDoc.Project = doc.Project;
        //        if (!String.IsNullOrEmpty(doc.Reference1)) oDoc.Reference1 = doc.Reference1;
        //        if (!String.IsNullOrEmpty(doc.Reference2)) oDoc.Reference2 = doc.Reference2;
        //        if (doc.RequriedDate != DateTime.MinValue) oDoc.RequriedDate = doc.RequriedDate;
        //        if (doc.Rounding) oDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES;
        //        if (doc.RoundingDiffAmount > 0) oDoc.RoundingDiffAmount = doc.RoundingDiffAmount;
        //        if (doc.SalesPersonCode > 0) oDoc.SalesPersonCode = doc.SalesPersonCode;
        //        if (!string.IsNullOrEmpty(doc.ShipToCode)) oDoc.ShipToCode = doc.ShipToCode;
        //        if (!String.IsNullOrEmpty(doc.TrackingNumber)) oDoc.TrackingNumber = doc.TrackingNumber;
        //        if (!String.IsNullOrEmpty(doc.U_GTPREFNO)) oDoc.UserFields.Fields.Item("U_GTPREFNO").Value = doc.U_GTPREFNO;
        //        #endregion

        //        #region Details
        //        if (doc.Lines != null)
        //        {
        //            for (int i = 0; i < doc.Lines.Count; i++)
        //            {

        //                if (i > 0) oDoc.Lines.Add();
        //                oDoc.Lines.SetCurrentLine(i);
        //                if (!String.IsNullOrEmpty(doc.Lines[i].AccountCode)) oDoc.Lines.AccountCode = doc.Lines[i].AccountCode;
        //                if (doc.Lines[i].AgreementNo > 0)
        //                {
        //                    oDoc.Lines.AgreementNo = doc.Lines[i].AgreementNo;
        //                    oDoc.Lines.AgreementRowNumber = doc.Lines[i].AgreementRowNumber;
        //                }
        //                if (doc.Lines[i].BackOrder) oDoc.Lines.BackOrder = SAPbobsCOM.BoYesNoEnum.tYES;
        //                if (doc.Lines[i].BaseEntry > 0)
        //                {
        //                    oDoc.Lines.BaseEntry = doc.Lines[i].BaseEntry;
        //                    oDoc.Lines.BaseLine = doc.Lines[i].BaseLine;
        //                    oDoc.Lines.BaseType = doc.Lines[i].BaseType;
        //                }
        //                if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode2)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode2;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode3)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode3;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode4)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode4;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode5)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode5;
        //                if (doc.Lines[i].DiscountPercent != 0) oDoc.Lines.DiscountPercent = doc.Lines[i].DiscountPercent;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].ItemCode)) oDoc.Lines.ItemCode = doc.Lines[i].ItemCode;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].ItemDescription)) oDoc.Lines.ItemDescription = doc.Lines[i].ItemDescription;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].ItemDetails)) oDoc.Lines.ItemDetails = doc.Lines[i].ItemDetails;
        //                if (doc.Lines[i].LineTotal > 0 && doc.DocType > 0) oDoc.Lines.LineTotal = doc.Lines[i].LineTotal;
        //                if (doc.Lines[i].Price > 0) oDoc.Lines.Price = doc.Lines[i].Price;
        //                if (doc.Lines[i].PriceAfterVAT > 0) oDoc.Lines.PriceAfterVAT = doc.Lines[i].PriceAfterVAT;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].ProjectCode)) oDoc.Lines.ProjectCode = doc.Lines[i].ProjectCode;
        //                if (doc.Lines[i].Quantity > 0) oDoc.Lines.Quantity = doc.Lines[i].Quantity;
        //                //if (!String.IsNullOrEmpty(doc.Lines[i].TaxCode)) oDoc.Lines.TaxCode = doc.Lines[i].TaxCode;
        //                //if (doc.Lines[i].TaxOnly) oDoc.Lines.TaxOnly = SAPbobsCOM.BoYesNoEnum.tYES;
        //                if (doc.Lines[i].TaxTotal > 0) oDoc.Lines.TaxTotal = doc.Lines[i].TaxTotal;
        //                if (doc.Lines[i].UnitPrice > 0) oDoc.Lines.UnitPrice = doc.Lines[i].UnitPrice;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].VatGroup)) oDoc.Lines.VatGroup = doc.Lines[i].VatGroup;
        //                if (doc.Lines[i].U_GOLD_PRICE != 0) oDoc.Lines.UserFields.Fields.Item("U_GOLD_PRICE").Value = doc.Lines[i].U_GOLD_PRICE;
        //                if (doc.Lines[i].U_GrossWeight != 0) oDoc.Lines.UserFields.Fields.Item("U_GrossWeight").Value = doc.Lines[i].U_GrossWeight;
        //                if (doc.Lines[i].U_PREMIUM != 0) oDoc.Lines.UserFields.Fields.Item("U_PREMIUM").Value = doc.Lines[i].U_PREMIUM;
        //                if (doc.Lines[i].U_PURITY != 0) oDoc.Lines.UserFields.Fields.Item("U_PURITY").Value = doc.Lines[i].U_PURITY;
        //                if (doc.Lines[i].U_REFINING_FEE != 0) oDoc.Lines.UserFields.Fields.Item("U_REFINING_FEE").Value = doc.Lines[i].U_REFINING_FEE;

        //            }
        //        }
        //        #endregion

        //        if (oDoc.Add() != 0)
        //        {
        //            errMsg = oCom.GetLastErrorDescription();
        //            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //            oDoc = null;
        //            return false;
        //        }
        //        oCom.GetNewObjectCode(out docentry);
        //        if (docentry == "")
        //        {
        //            errMsg = "Unknown Error! Please try again!";
        //            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //            oDoc = null;
        //            return false;
        //        }

        //        SAPbobsCOM.Documents getDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

        //        if (!getDoc.GetByKey(int.Parse(docentry)))
        //        {
        //            errMsg = "Unknown Error! Please try again!";
        //            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //            oDoc = null;
        //            return false;
        //        }
        //        else
        //        {
        //            if (!String.IsNullOrEmpty(doc.U_GTPREFNO))
        //            {
        //                if (getDoc.UserFields.Fields.Item("U_GTPREFNO").Value.ToString() != doc.U_GTPREFNO)
        //                {
        //                    errMsg = "Unknown Error! Please try again!";
        //                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //                    oDoc = null;
        //                    return false;
        //                }
        //            }
        //        }
        //        if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(getDoc);
        //        oDoc = null;
        //        getDoc = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        errMsg = ex.Message;
        //        return false;
        //    }
        //    return true;
        //}
        //public bool createAcePO(SAPB1WebApi.Ace.AceDoc doc, ref string docentry)
        //{
        //    if (!generateLog(doc, "22", "")) return false;

        //    try
        //    {
        //        oCom.StartTransaction();
        //        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);
        //        #region Header
        //        if (doc.HandWritten)
        //        {
        //            oDoc.DocNum = doc.DocNum;
        //            oDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES;
        //        }
        //        else
        //        {
        //            if (doc.Series > 0) oDoc.Series = doc.Series;
        //        }
        //        oDoc.DocDate = doc.DocDate;
        //        oDoc.CardCode = doc.CardCode;
        //        if (!String.IsNullOrEmpty(doc.ClosingRemarks)) oDoc.ClosingRemarks = doc.ClosingRemarks;
        //        if (!String.IsNullOrEmpty(doc.Comments)) oDoc.Comments = doc.Comments;
        //        if (doc.DiscountPercent != 0) oDoc.DiscountPercent = doc.DiscountPercent;
        //        if (!String.IsNullOrEmpty(doc.DocCurrency)) oDoc.DocCurrency = doc.DocCurrency;
        //        if (doc.DocDueDate != DateTime.MinValue) oDoc.DocDueDate = doc.DocDueDate;
        //        if (doc.DocRate > 0) oDoc.DocRate = doc.DocRate;
        //        if (doc.DocType > 0) oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service; else oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
        //        if (!String.IsNullOrEmpty(doc.JournalMemo)) oDoc.JournalMemo = doc.JournalMemo;
        //        if (!String.IsNullOrEmpty(doc.NumAtCard)) oDoc.NumAtCard = doc.NumAtCard;
        //        if (!String.IsNullOrEmpty(doc.OpeningRemarks)) oDoc.OpeningRemarks = doc.OpeningRemarks;
        //        if (!String.IsNullOrEmpty(doc.Project)) oDoc.Project = doc.Project;
        //        if (!String.IsNullOrEmpty(doc.Reference1)) oDoc.Reference1 = doc.Reference1;
        //        if (!String.IsNullOrEmpty(doc.Reference2)) oDoc.Reference2 = doc.Reference2;
        //        if (doc.RequriedDate != DateTime.MinValue) oDoc.RequriedDate = doc.RequriedDate;
        //        if (doc.Rounding) oDoc.Rounding = SAPbobsCOM.BoYesNoEnum.tYES;
        //        if (doc.RoundingDiffAmount > 0) oDoc.RoundingDiffAmount = doc.RoundingDiffAmount;
        //        if (doc.SalesPersonCode > 0) oDoc.SalesPersonCode = doc.SalesPersonCode;
        //        if (!string.IsNullOrEmpty(doc.ShipToCode)) oDoc.ShipToCode = doc.ShipToCode;
        //        if (!String.IsNullOrEmpty(doc.TrackingNumber)) oDoc.TrackingNumber = doc.TrackingNumber;
        //        if (!String.IsNullOrEmpty(doc.U_GTPREFNO)) oDoc.UserFields.Fields.Item("U_GTPREFNO").Value = doc.U_GTPREFNO;
        //        #endregion

        //        #region Details
        //        if (doc.Lines != null)
        //        {
        //            for (int i = 0; i < doc.Lines.Count; i++)
        //            {

        //                if (i > 0) oDoc.Lines.Add();
        //                oDoc.Lines.SetCurrentLine(i);
        //                if (!String.IsNullOrEmpty(doc.Lines[i].AccountCode)) oDoc.Lines.AccountCode = doc.Lines[i].AccountCode;
        //                if (doc.Lines[i].AgreementNo > 0)
        //                {
        //                    oDoc.Lines.AgreementNo = doc.Lines[i].AgreementNo;
        //                    oDoc.Lines.AgreementRowNumber = doc.Lines[i].AgreementRowNumber;
        //                }
        //                if (doc.Lines[i].BackOrder) oDoc.Lines.BackOrder = SAPbobsCOM.BoYesNoEnum.tYES;
        //                if (doc.Lines[i].BaseEntry > 0)
        //                {
        //                    oDoc.Lines.BaseEntry = doc.Lines[i].BaseEntry;
        //                    oDoc.Lines.BaseLine = doc.Lines[i].BaseLine;
        //                    oDoc.Lines.BaseType = doc.Lines[i].BaseType;
        //                }
        //                if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode2)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode2;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode3)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode3;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode4)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode4;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].CostingCode5)) oDoc.Lines.CostingCode = doc.Lines[i].CostingCode5;
        //                if (doc.Lines[i].DiscountPercent != 0) oDoc.Lines.DiscountPercent = doc.Lines[i].DiscountPercent;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].ItemCode)) oDoc.Lines.ItemCode = doc.Lines[i].ItemCode;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].ItemDescription)) oDoc.Lines.ItemDescription = doc.Lines[i].ItemDescription;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].ItemDetails)) oDoc.Lines.ItemDetails = doc.Lines[i].ItemDetails;
        //                if (doc.Lines[i].LineTotal > 0 && doc.DocType > 0) oDoc.Lines.LineTotal = doc.Lines[i].LineTotal;
        //                if (doc.Lines[i].Price > 0) oDoc.Lines.Price = doc.Lines[i].Price;
        //                if (doc.Lines[i].PriceAfterVAT > 0) oDoc.Lines.PriceAfterVAT = doc.Lines[i].PriceAfterVAT;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].ProjectCode)) oDoc.Lines.ProjectCode = doc.Lines[i].ProjectCode;
        //                if (doc.Lines[i].Quantity > 0) oDoc.Lines.Quantity = doc.Lines[i].Quantity;
        //                //if (!String.IsNullOrEmpty(doc.Lines[i].TaxCode)) oDoc.Lines.TaxCode = doc.Lines[i].TaxCode;
        //                //if (doc.Lines[i].TaxOnly) oDoc.Lines.TaxOnly = SAPbobsCOM.BoYesNoEnum.tYES;
        //                if (doc.Lines[i].TaxTotal > 0) oDoc.Lines.TaxTotal = doc.Lines[i].TaxTotal;
        //                if (doc.Lines[i].UnitPrice > 0) oDoc.Lines.UnitPrice = doc.Lines[i].UnitPrice;
        //                if (!String.IsNullOrEmpty(doc.Lines[i].VatGroup)) oDoc.Lines.VatGroup = doc.Lines[i].VatGroup;
        //                if (doc.Lines[i].U_GOLD_PRICE != 0) oDoc.Lines.UserFields.Fields.Item("U_GOLD_PRICE").Value = doc.Lines[i].U_GOLD_PRICE;
        //                if (doc.Lines[i].U_GrossWeight != 0) oDoc.Lines.UserFields.Fields.Item("U_GrossWeight").Value = doc.Lines[i].U_GrossWeight;
        //                if (doc.Lines[i].U_PREMIUM != 0) oDoc.Lines.UserFields.Fields.Item("U_PREMIUM").Value = doc.Lines[i].U_PREMIUM;
        //                if (doc.Lines[i].U_PURITY != 0) oDoc.Lines.UserFields.Fields.Item("U_PURITY").Value = doc.Lines[i].U_PURITY;
        //                if (doc.Lines[i].U_REFINING_FEE != 0) oDoc.Lines.UserFields.Fields.Item("U_REFINING_FEE").Value = doc.Lines[i].U_REFINING_FEE;

        //            }
        //        }
        //        #endregion

        //        if (oDoc.Add() != 0)
        //        {
        //            errMsg = oCom.GetLastErrorDescription();
        //            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //            oDoc = null;
        //            return false;
        //        }
        //        oCom.GetNewObjectCode(out docentry);
        //        if (docentry == "")
        //        {
        //            errMsg = "Unknown Error! Please try again!";
        //            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //            oDoc = null;
        //            return false;
        //        }

        //        SAPbobsCOM.Documents getDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders);

        //        if (!getDoc.GetByKey(int.Parse(docentry)))
        //        {
        //            errMsg = "Unknown Error! Please try again!";
        //            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //            oDoc = null;
        //            getDoc = null;
        //            return false;
        //        }
        //        else
        //        {
        //            if (!String.IsNullOrEmpty(doc.U_GTPREFNO))
        //            {
        //                if (getDoc.UserFields.Fields.Item("U_GTPREFNO").Value.ToString() != doc.U_GTPREFNO)
        //                {
        //                    errMsg = "Unknown Error! Please try again!";
        //                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //                    oDoc = null;
        //                    getDoc = null;
        //                    return false;
        //                }
        //            }
        //        }
        //        if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(getDoc);
        //        oDoc = null;
        //        getDoc = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        errMsg = ex.Message;
        //        return false;
        //    }
        //    return true;
        //}
        //public bool createAceGRN(SAPB1WebApi.Ace.AceGRN doc, ref string docentry)
        //{
        //    if (!generateLog(doc, "20", "")) return false;
        //    try
        //    {
        //        DataTable dtso = new DataTable();
        //        DataTable dtitems = new DataTable();
        //        using (SqlConnection conn = new SqlConnection(getConnectionString()))
        //        {
        //            using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
        //            {
        //                for (int i = 0; i < doc.SelectedPO.Count; i++)
        //                {
        //                    dat.SelectCommand.CommandText = "SELECT ISNULL(@Sequence,0) AS [Sequence],T0.DocNum, T0.DocEntry, T1.LineNum, T1.ItemCode, T1.Quantity, T1.OpenQty,  " +
        //                        "T1.Price, T1.U_GOLD_PRICE, T1.U_REFINING_FEE, T1.VatGroup, T1.WhsCode, ISNULL(T2.U_U_PTYPE,'') AS [PType],U_U_PAIR, T1.OpenQty AS [BalanceQty], " +
        //                        "ISNULL(T2.U_U_UQTY,0) AS [UniqueQty] " +
        //                        "FROM OPOR T0 INNER JOIN POR1 T1 ON T1.DocEntry = T0.DocEntry " +
        //                        "INNER JOIN OITM T2 ON T2.ItemCode = T1.ItemCode " +
        //                        "WHERE T0.CardCode = @CardCode AND T0.DocEntry = @DocEntry AND t1.LineNum = @LineNum ";
        //                    dat.SelectCommand.Parameters.Clear();
        //                    dat.SelectCommand.Parameters.AddWithValue("@CardCode", doc.CardCode);
        //                    dat.SelectCommand.Parameters.AddWithValue("@DocEntry", doc.SelectedPO[i].DocEntry);
        //                    dat.SelectCommand.Parameters.AddWithValue("@LineNum", doc.SelectedPO[i].LineNum);
        //                    dat.SelectCommand.Parameters.AddWithValue("@Sequence", doc.SelectedPO[i].Sequence);
        //                    dat.Fill(dtso);
        //                }
        //                for (int i = 0; i < doc.Items.Count; i++)
        //                {
        //                    dat.SelectCommand.CommandText = "SELECT T0.ItemCode, CASE WHEN ISNULL(@Quantity,0) = 0 THEN ROUND( ISNULL(@GrossWeight,0) * ISNULL(@Purity,0),T1.QtyDec) ELSE @Quantity END AS [Quantity], " +
        //                        "ISNULL(@GrossWeight,0) AS [GrossWeight], " +
        //                        "ISNULL(@Purity,0) AS [Purity], T0.ItemName, ISNULL(T0.U_U_SEQ,0) AS [Sequence], " +
        //                        "ISNULL(T0.U_U_PAIR,'') AS [PairItem], ISNULL(T0.U_U_PTYPE,'') AS [PType], ISNULL(T0.U_U_UQTY,0) AS [UniqueQty], " +
        //                        " CASE WHEN ISNULL(@Quantity,0) = 0 THEN ISNULL(@GrossWeight,0) * ISNULL(@Purity,0) ELSE @Quantity END AS [BalanceQty], " +
        //                        " T1.QtyDec, T0.ManBtchNum " +
        //                        "FROM OITM T0  " +
        //                        "LEFT OUTER JOIN OADM T1 ON 0 = 0 " +
        //                        "WHERE T0.ItemCode = @ItemCode ";
        //                    dat.SelectCommand.Parameters.Clear();
        //                    dat.SelectCommand.Parameters.AddWithValue("@ItemCode", doc.Items[i].itemCode);
        //                    dat.SelectCommand.Parameters.AddWithValue("@Quantity", doc.Items[i].Quantity);
        //                    dat.SelectCommand.Parameters.AddWithValue("@GrossWeight", doc.Items[i].U_GrossWeight);
        //                    dat.SelectCommand.Parameters.AddWithValue("@Purity", doc.Items[i].U_PURITY);
        //                    dat.Fill(dtitems);
        //                }
        //            }
        //        }

        //        if (dtso.Rows.Count <= 0) throw new Exception("Selected PO does not found!");
        //        if (dtitems.Rows.Count <= 0) throw new Exception("Items is empty!");

        //        DataTable dtso1 = new DataTable();
        //        dtso1.Columns.Add("DocEntry", typeof(int));
        //        dtso1.Columns.Add("LineNum", typeof(int));
        //        dtso1.Columns.Add("ItemCode");
        //        dtso1.Columns.Add("Quantity", typeof(decimal));
        //        dtso1.Columns.Add("Purity", typeof(decimal));
        //        dtso1.Columns.Add("ManBtchNum");

        //        DataRow[] sorows = dtso.Select(null, "Sequence ASC");

        //        #region assing item by PO sequence
        //        foreach (DataRow row in sorows)
        //        {
        //            string itemCode = "";
        //            string nsitem = row["ItemCode"].ToString();
        //            decimal openqty = decimal.Parse(row["OpenQty"].ToString());
        //            int entryno = int.Parse(row["DocEntry"].ToString());
        //            int linenum = int.Parse(row["LineNum"].ToString());
        //            if (openqty <= 0) continue;
        //            decimal appliedqty = 0, itemqty = 0;
        //            int rounqty = 3;
        //            string paringType = row["PType"].ToString();

        //            string filter = "PType = '" + paringType + "' AND BalanceQty > 0 ";
        //            if (paringType == "P") filter += " AND PairItem = '" + nsitem + "' ";

        //            foreach (DataRow dr in dtitems.Select(filter))
        //            {
        //                decimal uniqueqty = decimal.Parse(dr["UniqueQty"].ToString());
        //                itemqty = decimal.Parse(dr["BalanceQty"].ToString());
        //                rounqty = int.Parse(dr["QtyDec"].ToString());
        //                itemqty = Math.Round(itemqty, rounqty, MidpointRounding.AwayFromZero);
        //                itemCode = dr["ItemCode"].ToString();
        //                if (uniqueqty > 0)
        //                {
        //                    itemqty = Math.Floor(itemqty / uniqueqty) * uniqueqty;
        //                }
        //                if (itemqty <= 0) continue;
        //                if (openqty >= itemqty)
        //                {
        //                    if (uniqueqty > 0)
        //                    {
        //                        appliedqty = Math.Floor(itemqty / uniqueqty) * uniqueqty;
        //                    }
        //                    else
        //                    {
        //                        appliedqty = itemqty;
        //                    }
        //                }
        //                else
        //                {
        //                    if (uniqueqty > 0)
        //                    {
        //                        appliedqty = Math.Floor(openqty / uniqueqty) * uniqueqty;
        //                    }
        //                    else
        //                    {
        //                        appliedqty = openqty;
        //                    }
        //                }
        //                openqty = openqty - appliedqty;
        //                if (appliedqty <= 0) continue;

        //                dtso1.Rows.Add(entryno, linenum, itemCode, appliedqty, dr["Purity"], dr["ManBtchNum"]);
        //                dr["BalanceQty"] = itemqty - appliedqty;
        //                if (openqty <= 0) break;
        //            }
        //            row["BalanceQty"] = openqty;

        //        }
        //        #endregion

        //        #region assign balance item

        //        foreach (DataRow row in dtitems.Select("BalanceQty > 0", "UniqueQty DESC"))
        //        {
        //            decimal uniqueqty = decimal.Parse(row["UniqueQty"].ToString());
        //            decimal itemqty = decimal.Parse(row["BalanceQty"].ToString());
        //            int rounqty = int.Parse(row["QtyDec"].ToString());
        //            itemqty = Math.Round(itemqty, rounqty, MidpointRounding.AwayFromZero);
        //            string itemCode = row["ItemCode"].ToString();
        //            string pType = row["PType"].ToString();
        //            string filter = "BalanceQty > 0";
        //            if (uniqueqty > 0) filter = "BalanceQty > " + uniqueqty.ToString();
        //            foreach (DataRow dr in dtso.Select(filter, "Sequence ASC"))
        //            {
        //                decimal openqty = decimal.Parse(dr["BalanceQty"].ToString());
        //                decimal appliedqty = 0;
        //                if (uniqueqty > 0)
        //                {
        //                    if (openqty >= itemqty)
        //                    {
        //                        appliedqty = Math.Floor(itemqty / uniqueqty) * uniqueqty;
        //                    }
        //                    else
        //                    {
        //                        appliedqty = Math.Floor(openqty / uniqueqty) * uniqueqty;
        //                    }
        //                }
        //                else
        //                {
        //                    if (openqty >= itemqty)
        //                    {
        //                        appliedqty = itemqty;
        //                    }
        //                    else
        //                    {
        //                        appliedqty = openqty;
        //                    }
        //                }
        //                if (appliedqty <= 0) continue;
        //                itemqty = itemqty - appliedqty;
        //                dr["BalanceQty"] = openqty - appliedqty;
        //                dtso1.Rows.Add(dr["DocEntry"], dr["LineNum"], itemCode, appliedqty, row["Purity"], row["ManBtchNum"]);
        //                if (itemqty <= 0) break;
        //            }
        //            row["BalanceQty"] = itemqty;
        //        }
        //        #endregion

        //        #region assign over grn item for unique qty
        //        foreach (DataRow row in dtitems.Select("BalanceQty > 0 AND UniqueQty > 0", "UniqueQty DESC"))
        //        {
        //            decimal uniqueqty = decimal.Parse(row["UniqueQty"].ToString());
        //            decimal itemqty = decimal.Parse(row["BalanceQty"].ToString());
        //            int rounqty = int.Parse(row["QtyDec"].ToString());
        //            itemqty = Math.Round(itemqty, rounqty, MidpointRounding.AwayFromZero);
        //            string itemCode = row["ItemCode"].ToString();
        //            string pairItem = row["PairItem"].ToString();

        //            DataRow[] drs;
        //            if (itemqty < uniqueqty)
        //            {
        //                continue;
        //                //drs = dtso.Select("BalanceQty > 0 AND PType in ('S','R')", "Sequence ASC");
        //                //if (drs.Length <= 0) drs = dtso.Select("PType in ('S','R')", "Sequence ASC");

        //            }
        //            else
        //            {
        //                drs = dtso.Select("BalanceQty > 0 AND ItemCode = '" + pairItem + "' ", "Sequence ASC");
        //                if (drs.Length <= 0) drs = dtso.Select("ItemCode = '" + pairItem + "' ", "Sequence ASC");
        //                if (drs.Length <= 0) drs = dtso.Select("BalanceQty > 0 AND PType <> 'P' ", "Sequence ASC");
        //                if (drs.Length <= 0) drs = dtso.Select("PType <> 'P' ", "Sequence ASC");
        //                if (drs.Length <= 0) drs = dtso.Select("", "Sequence ASC");
        //            }
        //            DataRow dr;
        //            if (drs.Length <= 0) continue;
        //            else dr = drs[0];

        //            decimal appliedqty = 0;

        //            if (uniqueqty > 0)
        //            {
        //                appliedqty = Math.Floor(itemqty / uniqueqty) * uniqueqty;
        //            }
        //            else appliedqty = itemqty;
        //            dtso1.Rows.Add(dr["DocEntry"], dr["LineNum"], itemCode, appliedqty, row["Purity"], row["ManBtchNum"]);

        //        }
        //        #endregion

        //        foreach (DataRow row in dtso.Rows)
        //        {
        //            string entryno = row["DocEntry"].ToString();
        //            string lineNum = row["LineNum"].ToString();
        //            decimal totalqty = 0;
        //            foreach (DataRow dr in dtso1.Select("DocEntry = " + entryno + " AND LineNum = " + lineNum))
        //            {
        //                decimal qty = decimal.Parse(dr["Quantity"].ToString());
        //                totalqty += qty;
        //            }
        //            row["Quantity"] = totalqty;
        //        }

        //        oCom.StartTransaction();
        //        SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
        //        #region Header
        //        oDoc.DocObjectCode = SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes;
        //        if (doc.HandWritten)
        //        {
        //            oDoc.DocNum = doc.DocNum;
        //            oDoc.HandWritten = SAPbobsCOM.BoYesNoEnum.tYES;
        //        }
        //        else
        //        {
        //            if (doc.Series > 0) oDoc.Series = doc.Series;
        //        }
        //        oDoc.DocDate = doc.DocDate;
        //        oDoc.CardCode = doc.CardCode;
        //        if (doc.DocDueDate != DateTime.MinValue) oDoc.DocDueDate = doc.DocDueDate;
        //        oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
        //        if (!String.IsNullOrEmpty(doc.NumAtCard)) oDoc.NumAtCard = doc.NumAtCard;
        //        if (!String.IsNullOrEmpty(doc.U_GTPREFNO)) oDoc.UserFields.Fields.Item("U_GTPREFNO").Value = doc.U_GTPREFNO;
        //        #endregion

        //        #region Details
        //        int line = -1;
        //        foreach (DataRow row in dtso.Rows)
        //        {
        //            double qty = double.Parse(row["Quantity"].ToString());
        //            if (qty == 0) continue;

        //            line++;
        //            int entryno = int.Parse(row["DocEntry"].ToString());
        //            int linenum = int.Parse(row["LineNum"].ToString());

        //            if (line > 0) oDoc.Lines.Add();
        //            oDoc.Lines.SetCurrentLine(line);
        //            oDoc.Lines.ItemCode = row["ItemCode"].ToString();
        //            oDoc.Lines.Quantity = double.Parse(row["Quantity"].ToString());
        //            oDoc.Lines.UnitPrice = 0;
        //            oDoc.Lines.BaseEntry = entryno;
        //            oDoc.Lines.BaseLine = linenum;
        //            oDoc.Lines.BaseType = 22;
        //            oDoc.Lines.UserFields.Fields.Item("U_PO_GRP").Value = int.Parse(row["DocNum"].ToString());

        //            foreach (DataRow dr in dtso1.Select("DocEntry = " + entryno.ToString() + " AND LineNum = " + linenum.ToString()))
        //            {
        //                line++;
        //                oDoc.Lines.Add();
        //                oDoc.Lines.SetCurrentLine(line);
        //                oDoc.Lines.ItemCode = dr["ItemCode"].ToString();
        //                oDoc.Lines.Quantity = double.Parse(dr["Quantity"].ToString());
        //                oDoc.Lines.UnitPrice = double.Parse(row["Price"].ToString());
        //                oDoc.Lines.WarehouseCode = row["WhsCode"].ToString();
        //                oDoc.Lines.VatGroup = row["VatGroup"].ToString();
        //                oDoc.Lines.UserFields.Fields.Item("U_PURITY").Value = double.Parse(dr["Purity"].ToString());
        //                oDoc.Lines.UserFields.Fields.Item("U_PO_GRP").Value = int.Parse(row["DocNum"].ToString());
        //                if (dr["ManBtchNum"].ToString() == "Y")
        //                {
        //                    oDoc.Lines.BatchNumbers.SetCurrentLine(0);
        //                    oDoc.Lines.BatchNumbers.ManufacturerSerialNumber = oDoc.CardCode;
        //                    oDoc.Lines.BatchNumbers.BatchNumber = String.Format("{0}-{1}", oDoc.CardCode, DateTime.Now.ToString("yyyyMMdd-HHmm"));
        //                    oDoc.Lines.BatchNumbers.UserFields.Fields.Item("U_Purity").Value = double.Parse(dr["Purity"].ToString());
        //                    oDoc.Lines.BatchNumbers.Quantity = double.Parse(dr["Quantity"].ToString());
        //                    oDoc.Lines.BatchNumbers.InternalSerialNumber = dr["ItemCode"].ToString();

        //                }

        //            }

        //        }
        //        #endregion

        //        if (oDoc.Add() != 0)
        //        {
        //            errMsg = oCom.GetLastErrorDescription();
        //            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //            oDoc = null;
        //            return false;
        //        }
        //        oCom.GetNewObjectCode(out docentry);
        //        if (docentry == "")
        //        {
        //            errMsg = "Unknown Error! Please try again!";
        //            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //            oDoc = null;
        //            return false;
        //        }

        //        SAPbobsCOM.Documents getDoc = (SAPbobsCOM.Documents)oCom.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);

        //        if (!getDoc.GetByKey(int.Parse(docentry)))
        //        {
        //            errMsg = "Unknown Error! Please try again!";
        //            if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //            oDoc = null;
        //            getDoc = null;
        //            return false;
        //        }
        //        else
        //        {
        //            if (!String.IsNullOrEmpty(doc.U_GTPREFNO))
        //            {
        //                if (getDoc.UserFields.Fields.Item("U_GTPREFNO").Value.ToString() != doc.U_GTPREFNO)
        //                {
        //                    errMsg = "Unknown Error! Please try again!";
        //                    if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //                    oDoc = null;
        //                    getDoc = null;
        //                    return false;
        //                }
        //            }
        //        }
        //        if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(getDoc);
        //        oDoc = null;
        //        getDoc = null;
        //        dtso.Dispose();
        //        dtso1.Dispose();
        //        dtitems.Dispose();
        //    }
        //    catch (Exception ex)
        //    {
        //        if (oCom.InTransaction) oCom.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        errMsg = ex.Message;
        //        return false;
        //    }
        //    return true;
        //}
        public bool getOpenPOList(ref DataSet ds, string blockDraft, string cardcode)
        {
            if (String.IsNullOrEmpty(cardcode)) cardcode = "";
            if (String.IsNullOrEmpty(blockDraft)) blockDraft = "";
            try
            {
                using (SqlConnection conn = new SqlConnection(getConnectionString()))
                {
                    using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
                    {
                        DataTable dt = new DataTable("opnlist");

                        string selectcmd = "SELECT T0.DocNum, T0.DocDate, T0.CardCode, T0.DocEntry, T1.LineNum, T1.ItemCode, T1.Dscription, T1.Quantity, T1.OpenQty, ISNULL(T2.Quantity,0) AS [DraftQty], " +
                            "T1.OpenQty - ISNULL(T2.Quantity,0) AS [opndraft], T1.Price, T0.DocTotal, T0.VatSum, T0.DocTotal - T0.VatSum AS [DocTotalAmt], " +
                            "CAST(CASE WHEN ISNULL(T2.Quantity,0) > 0 THEN 1 else 0 end AS BIT) as [draftGRN] " +
                            "FROM OPOR T0 " +
                            "INNER JOIN POR1 T1 ON T1.DocEntry = T0.DocEntry " +
                            "LEFT OUTER JOIN DRF1 T2 " +
                            "INNER JOIN ODRF T3 ON T3.DocEntry = T2.DocEntry AND T3.DocStatus = 'O' AND T3.ObjType = '20' " +
                            "ON T2.BaseEntry = T1.DocEntry AND T2.BaseLine = T1.LineNum AND T2.BaseType = T0.ObjType " +
                            "WHERE T0.DocStatus = 'O' AND T0.CANCELED = 'N' " +
                            "AND T1.OpenQty > 0 AND T0.CardCode = @CardCode ";

                        dat.SelectCommand.CommandText = selectcmd;
                        dat.SelectCommand.Parameters.Clear();
                        if (selectcmd.Contains("@CardCode")) dat.SelectCommand.Parameters.AddWithValue("@CardCode", cardcode);
                        dat.Fill(dt);

                        if (blockDraft.ToUpper() == "Y")
                        {
                            if (dt.Select("DraftQty > 0").Length > 0)
                            {
                                throw new Exception("Open Draft Exist!");
                            }
                        }

                        ds.Tables.Add(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return false;
            }
            return true;
        }

        public bool getAceBP(ref DataSet ds, string cardType, string cardCode)
        {
            if (String.IsNullOrEmpty(cardType)) cardType = "";
            if (String.IsNullOrEmpty(cardCode)) cardCode = "";
            try
            {
                using (SqlConnection conn = new SqlConnection(getConnectionString()))
                {
                    using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
                    {
                        DataTable dt = new DataTable("OCRD");

                        string selectcmd = "SELECT CardCode,CardName,CardType,frozenFor,AliasName FROM OCRD WHERE 0 = 0 ";
                        if (cardType != "") selectcmd += " AND CardType = @CardType ";
                        if (cardCode != "")
                        {
                            selectcmd += " AND CardCode = @CardCode ";
                        }

                        dat.SelectCommand.CommandText = selectcmd;
                        dat.SelectCommand.Parameters.Clear();
                        if (selectcmd.Contains("@CardType")) dat.SelectCommand.Parameters.AddWithValue("@CardType", cardType);
                        if (selectcmd.Contains("@CardCode")) dat.SelectCommand.Parameters.AddWithValue("@CardCode", cardCode);
                        dat.Fill(dt);

                        ds.Tables.Add(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return false;
            }
            return true;
        }

        public bool getAceRateCard(ref DataSet ds, string cardcode, string itemcode)
        {
            if (String.IsNullOrEmpty(cardcode)) cardcode = "";
            if (String.IsNullOrEmpty(itemcode)) itemcode = "";
            try
            {
                using (SqlConnection conn = new SqlConnection(getConnectionString()))
                {
                    using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
                    {
                        DataTable dt = new DataTable("ratecard");

                        string selectcmd = "SELECT u_cardcode, u_itemcode, u_purity FROM [@BP_RATETABLE] WHERE 0 = 0 ";
                        if (cardcode != "") selectcmd += " AND U_CardCode = @cardcode ";
                        if (itemcode != "")
                        {
                            selectcmd += " AND U_ItemCode = @itemcode ";
                        }

                        dat.SelectCommand.CommandText = selectcmd;
                        dat.SelectCommand.Parameters.Clear();
                        if (selectcmd.Contains("@cardcode")) dat.SelectCommand.Parameters.AddWithValue("@CardType", cardcode);
                        if (selectcmd.Contains("@itemcode")) dat.SelectCommand.Parameters.AddWithValue("@itemcode", itemcode);
                        dat.Fill(dt);

                        ds.Tables.Add(dt);
                    }
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return false;
            }
            return true;
        }

        public bool getAceDocument(ref DataSet ds, string doctype, string keytype, string key)
        {
            if (String.IsNullOrEmpty(key)) key = "";
            if (String.IsNullOrEmpty(keytype)) keytype = "";
            string hdrtbl = "", dtltbl = "", colname = "DocEntry";
            switch (doctype)
            {
                case "22":
                    hdrtbl = "OPOR";
                    dtltbl = "POR1";
                    break;
                case "17":
                    hdrtbl = "ORDR";
                    dtltbl = "RDR1";
                    break;
                case "20":
                    hdrtbl = "ODRF";
                    dtltbl = "DRF1";
                    break;
            }
            if (keytype.ToUpper() == "REF") colname = "U_GTPREFNO";
            try
            {
                using (SqlConnection conn = new SqlConnection(getConnectionString()))
                {
                    using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
                    {
                        DataTable dt = new DataTable("hdr");
                        DataTable dt1 = new DataTable("dtl");

                        dat.SelectCommand.CommandText = "SELECT * FROM " + hdrtbl + " WHERE " + colname + " = @col AND ObjType = @doctype ";
                        dat.SelectCommand.Parameters.Clear();
                        dat.SelectCommand.Parameters.AddWithValue("@doctype", doctype);
                        dat.SelectCommand.Parameters.AddWithValue("@col", key);
                        dat.Fill(dt);

                        string docentry = "0";
                        if (dt.Rows.Count > 0) docentry = dt.Rows[0]["DocEntry"].ToString();

                        dat.SelectCommand.CommandText = "SELECT * FROM " + dtltbl + " WHERE DocEntry = @docentry ";
                        dat.SelectCommand.Parameters.Clear();
                        dat.SelectCommand.Parameters.AddWithValue("@docentry", docentry);
                        dat.Fill(dt1);

                        ds.Tables.Add(dt);
                        ds.Tables.Add(dt1);
                    }
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return false;
            }
            return true;
        }

        public bool getAceItems(ref DataSet ds, string whscode, string itemcode)
        {
            if (String.IsNullOrEmpty(whscode)) whscode = "";
            if (String.IsNullOrEmpty(itemcode)) itemcode = "";
            try
            {
                using (SqlConnection conn = new SqlConnection(getConnectionString()))
                {
                    using (SqlDataAdapter dat = new SqlDataAdapter("", conn))
                    {
                        DataTable dt = new DataTable("oitm");
                        DataTable dt1 = new DataTable("oitw");
                        string selectcmd = "SELECT ItemCode, ItemName, FrgnName FROM OITM WHERE FrozenFor = 'N' ";
                        if (itemcode != "") selectcmd += " AND ItemCode = @ItemCode ";

                        dat.SelectCommand.CommandText = selectcmd;
                        dat.SelectCommand.Parameters.Clear();
                        if (selectcmd.Contains("@ItemCode")) dat.SelectCommand.Parameters.AddWithValue("@ItemCode", itemcode);
                        dat.Fill(dt);



                        selectcmd = "SELECT * FROM OITW WHERE 0 = 0 ";
                        if (itemcode != "") selectcmd += " AND ItemCode = @ItemCode ";
                        if (whscode != "")
                        {
                            if (whscode.ToUpper() != "ALL")
                            {
                                selectcmd += " AND WhsCode = @WhsCode ";
                            }
                        }
                        dat.SelectCommand.CommandText = selectcmd;
                        dat.SelectCommand.Parameters.Clear();
                        if (selectcmd.Contains("@ItemCode")) dat.SelectCommand.Parameters.AddWithValue("@ItemCode", itemcode);
                        if (selectcmd.Contains("@WhsCode")) dat.SelectCommand.Parameters.AddWithValue("@WhsCode", whscode);
                        dat.Fill(dt1);

                        ds.Tables.Add(dt);
                        ds.Tables.Add(dt1);
                    }
                }
            }
            catch (Exception ex)
            {
                errMsg = ex.Message;
                return false;
            }
            return true;
        }
    }
}
