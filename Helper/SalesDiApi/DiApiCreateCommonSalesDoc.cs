using Dapper;
using Microsoft.Data.SqlClient;
using IMAppSapMidware_NetCore.Models.CommonSalesDocs;
using IMAppSapMidware_NetCore.Models.Share;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IMAppSapMidware_NetCore.Helper.DiApi
{
    public class DiApiCreateCommonSalesDoc
    {
        public ErpPropertyHelper ErpProperty { get; set; } = null;
        public string PostedDocNum { get; set; } = string.Empty;
        public string LastSAPMsg { get; set; } = string.Empty;
        public string Midware_DBConnStr { get; set; } = string.Empty;
        public string Erp_DBConnStr { get; set; } = string.Empty;
        public string Guid { get; set; } = string.Empty;
        public string DocType { get; set; } = string.Empty;
        public int AttachedFileCnt { get; set; } = -1;
        public DocHeader Doc { get; set; } = null;
        public List<DocDetail> DocDetails { get; set; } = null;
        private string curSapTableName { get; set; } = string.Empty;
        private Company sapCompany { get; set; } = null;
        private BoObjectTypes currentDocType { get; set; }

        //private readonly ILogger<Worker> Log;

        /// <summary>
        /// Constructor
        /// </summary>
        //public DiApiCreateCommonSalesDoc(ILogger<Worker> logger) => Log = logger;

        void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }

        public bool CreateCommDocument()
        {
            string modName = $"[CreateCommDocu  ment][{DocType}]";
            try
            {
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

                bool isPostDraft = Doc.IsPostDraft; // (ut.SafeGetInt(docHeader.Rows[0]["isPostDraft"]) == 1);
                Documents sapDoc = GetSAPDocType(DocType, isPostDraft);

                // user define field
                string docNUmber = Doc.DocNum; // ut.SafeGetStr(docHeader.Rows[0]["docNum"]);
                sapDoc.UserFields.Fields.Item("U_Guid").Value = $"{Doc.Guid}";  //guidVal;
                sapDoc.UserFields.Fields.Item("U_PhoneDocNumber").Value = docNUmber;


                // prepare the document header
                string cardCode = Doc.CardCode; //ut.SafeGetStr(docHeader.Rows[0]["cardCode"]);
                sapDoc.CardCode = cardCode;
                sapDoc.DocDate = Doc.DocDate; //ut.SafeGetDateTime(docHeader.Rows[0]["docDate"]);
                sapDoc.TaxDate = Doc.TaxDate;   //ut.SafeGetDateTime(docHeader.Rows[0]["taxDate"]);
                sapDoc.DocDueDate = Doc.DueDate; // ut.SafeGetDateTime(docHeader.Rows[0]["dueDate"]);

                // doc reference 2 -> special cut off the extra lenght
                if (!string.IsNullOrWhiteSpace(Doc.Ref2))
                {
                    sapDoc.Reference2 = Doc.Ref2.Substring(0, 11);
                }
                
                sapDoc.Comments = $"{Doc.Comments}"; //ut.SafeGetStr(docHeader.Rows[0]["comments"]);
                sapDoc.JournalMemo = $"{Doc.JrnlMemo}";// ut.SafeGetStr(docHeader.Rows[0]["jrnlMemo"]);
                sapDoc.NumAtCard = $"{Doc.NumberAtCard}";// ut.SafeGetStr(docHeader.Rows[0]["numberAtCard"]);

                var contactPersonName = Doc.ContactPerson;// ut.SafeGetStr(docHeader.Rows[0]["contactPerson"]);
                var resultContanctPersonId = GetContactPersonCode(cardCode, contactPersonName);
                if (resultContanctPersonId > -1)
                {
                    sapDoc.ContactPersonCode = resultContanctPersonId;
                }

                if (!string.IsNullOrWhiteSpace(Doc.ShipAddress))
                {
                    sapDoc.ShipToCode = Doc.ShipAddress; // ut.SafeGetStr(docHeader.Rows[0]["shipAddress"]);
                }

                if (!string.IsNullOrWhiteSpace(Doc.BillAddress))
                {
                    sapDoc.PayToCode = Doc.BillAddress; // ut.SafeGetStr(docHeader.Rows[0]["shipAddress"]);
                }

                //sapDoc.PayToCode = Doc.BillAddress; //ut.SafeGetStr(docHeader.Rows[0]["billAddress"]);
                sapDoc.DiscountPercent = Doc.DiscountByPercent; // // ut.SafeGetDouble(docHeader.Rows[0]["discountByPercent"]);

                // tax, currency will auto assign by sap auto
                // sales person code will follow the sap auto

                sapDoc.HandWritten = BoYesNoEnum.tNO;
                int updateSuccessCnt = 0;
                if (DocDetails != null)
                {
                    DocDetails.ForEach(line =>
                    {
                        sapDoc.Lines.ItemCode = line.ItemCode; //  ut.SafeGetStr(datarow["itemCode"]);  // sqItem[x].ItemCode;
                        sapDoc.Lines.Quantity = line.OrderQty; // ut.SafeGetDouble(datarow["orderQty"]); // sqItem[x].OrderQty;
                        sapDoc.Lines.UnitPrice = line.Price; // ut.SafeGetDouble(datarow["price"]); // sqItem[x].OrderQty;
                        sapDoc.Lines.DiscountPercent = line.DisByPercent;  //ut.SafeGetDouble(datarow["disByPercent"]);
                        sapDoc.Lines.TaxCode = line.TaxCode; //ut.SafeGetStr(datarow["taxCode"]);

                        var whsCode = GetWarehouseCode(line.Warehouse);
                        if (!string.IsNullOrWhiteSpace(whsCode))
                        {
                            sapDoc.Lines.WarehouseCode = whsCode; //ut.SafeGetStr(datarow["whsCode"]);
                        }
                        
                        sapDoc.Lines.ActualDeliveryDate = line.DeliveryDate; //ut.SafeGetDateTime(datarow["deliveryDate"]);
                        sapDoc.Lines.Add();
                        updateSuccessCnt++;
                    });
                }

                if (updateSuccessCnt != DocDetails.Count) // check does all record insert
                {
                  
                    Log($"{modName}\nSome items no update into the list, " +
                                 $"Marketing doc no created & roll back, Please try again later\n");
                    return false;
                }

                int addResult = sapDoc.Add();
                if (addResult != 0)
                {                    
                    Log($"{modName}\n{sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}\n");
                    return false;
                }

                // when add result = 0
                int newKey = Convert.ToInt32(sapCompany.GetNewObjectKey());
                string result = newKey.ToString(); // docentry from the object 

                int doNum = GetDocNumberbyDoEntry(result, curSapTableName);
                PostedDocNum = doNum.ToString();
                Log($"{modName} Created, DocEntry: {result} ,DocNo: {PostedDocNum}, {Doc.Guid}");

                // added 27-Dec-2019
                // for file attachment
                if (AttachedFileCnt == 0)
                {
                    return (!string.IsNullOrEmpty(PostedDocNum));
                }

                AddFileAttachment(newKey, Guid, currentDocType);    
                return (!string.IsNullOrEmpty(PostedDocNum));
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return false;
            }
        }

        void AddFileAttachment(int docEntry, string guidHeader, BoObjectTypes objectTypes)
        {
            try
            {
                var files = GetAttachmentFiles(guidHeader);
                if (files == null) return;
                if (files.Count == 0) return;

                Documents doc = (Documents)sapCompany.GetBusinessObject(objectTypes);
                doc.GetByKey(docEntry);

                var attachedments = (Attachments2)sapCompany.GetBusinessObject(BoObjectTypes.oAttachments2);
                int fileCnt = 0;

                files.ForEach(file =>
                {
                    if (File.Exists(file.ServerSavedPath))
                    {
                        attachedments.Lines.Add();
                        attachedments.Lines.FileName = Path.GetFileNameWithoutExtension(file.ServerSavedPath);
                        attachedments.Lines.FileExtension = Path.GetExtension(file.ServerSavedPath).Substring(1);
                        attachedments.Lines.SourcePath = Path.GetDirectoryName(file.ServerSavedPath);
                        attachedments.Lines.Override = BoYesNoEnum.tYES;
                        fileCnt++;
                    }
                });

                if (fileCnt != files.Count)
                {                    
                    Log($"Marketing Doc Attachments {sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
                    return;
                }

                if (attachedments.Add() == 0)
                {
                    int iAttEntry = int.Parse(sapCompany.GetNewObjectKey());
                    doc.AttachmentEntry = iAttEntry;
                    doc.Update();
                    Log($"{files.Count} File(s) Attached, DocNo: {PostedDocNum}\n{Doc.Guid}");
                    return;
                }

                // else                                    
                Log($"Marketing Doc Attachments {sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");                
            }
        }

        Documents GetSAPDocType(string docType, bool isPostDraft)
        {
            Documents retDoc = null;
            switch (docType)
            {
                case "Quotation":
                    {
                        if (isPostDraft)
                        {
                            retDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                            retDoc.DocObjectCode = BoObjectTypes.oQuotations;
                            retDoc.DocObjectCodeEx = "23"; // sales quotation type
                            curSapTableName = "ODRF";
                            currentDocType = BoObjectTypes.oDrafts;
                            return retDoc;
                        }
                        // ELSE
                        retDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oQuotations);
                        curSapTableName = "OQUT";
                        currentDocType = BoObjectTypes.oQuotations;
                        return retDoc;
                    }
                case "Sales Order":
                    {
                        if (isPostDraft)
                        {
                            retDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                            retDoc.DocObjectCode = BoObjectTypes.oOrders;
                            retDoc.DocObjectCodeEx = "17"; // sales order in draft type
                            curSapTableName = "ODRF";
                            currentDocType = BoObjectTypes.oDrafts;
                            return retDoc;
                        }
                        // ELSE
                        retDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oOrders);
                        curSapTableName = "ORDR";
                        currentDocType = BoObjectTypes.oOrders;
                        return retDoc;
                    }
                case "Invoice":
                    {
                        retDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                        retDoc.DocObjectCode = BoObjectTypes.oInvoices;
                        retDoc.DocObjectCodeEx = "13"; // sales quotation type
                        curSapTableName = "ODRF";
                        currentDocType = BoObjectTypes.oDrafts;
                        return retDoc;
                    }
                case "Delivery Order":
                    {
                        retDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                        retDoc.DocObjectCode = BoObjectTypes.oDeliveryNotes;
                        retDoc.DocObjectCodeEx = "15"; // Delivery order object type  
                        curSapTableName = "ODRF";
                        currentDocType = BoObjectTypes.oDrafts;
                        return retDoc;
                    }
                case "Credit Note":
                    {
                        retDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                        retDoc.DocObjectCode = BoObjectTypes.oCreditNotes;
                        retDoc.DocObjectCodeEx = "14"; // AR Credit Memo object type     
                        curSapTableName = "ODRF";
                        currentDocType = BoObjectTypes.oDrafts;
                        return retDoc;
                    }
                case "Return":
                    {
                        retDoc = (Documents)sapCompany.GetBusinessObject(BoObjectTypes.oDrafts);
                        retDoc.DocObjectCode = BoObjectTypes.oReturns;
                        retDoc.DocObjectCodeEx = "16"; // AR Return object type (Sales)
                        curSapTableName = "ODRF";
                        currentDocType = BoObjectTypes.oDrafts;
                        return retDoc;
                    }
                default:
                    {
                        return null;
                    }
            }
        }

        List<FileUpload> GetAttachmentFiles(string guid) // query from midware
        {
            try
            {
                var sqlQuery = @"SELECT * 
                                 FROM zmwFileUpload 
                                 WHERE HeaderGuid=@guid";

                var results = new SqlConnection(Midware_DBConnStr).Query<FileUpload>(sqlQuery, new { guid }).ToList();
                return results;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return null;
            }
        }

        string GetWarehouseCode(string warehoueName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(warehoueName))
                {
                    return string.Empty;
                }

                string sql = $"SELECT WhsCode" +
                            $" FROM OWHS" +
                            $" WHERE WhsName = @warehoueName";

                return new SqlConnection(Erp_DBConnStr).ExecuteScalar<string>(sql, new { warehoueName });
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return string.Empty;
            }
        }


        int GetContactPersonCode(string cardCode, string contactName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(contactName))
                {
                    return -1;
                }

                string sql = $"SELECT cntctCode" +
                            $" FROM OCPR" +
                            $" WHERE CardCode = @cardCode" +
                            $" AND Name = @contactName ";

                return new SqlConnection(Erp_DBConnStr).ExecuteScalar<int>(sql, new { cardCode, contactName });
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return -1;
            }
        }



        int GetDocNumberbyDoEntry(string docEntry, string tableName)
        {
            try
            {
                string sql = $@"SELECT DocNum  
                               FROM {tableName} 
                               WHERE DocEntry = @docEntry";

                var result = (int)new SqlConnection(Erp_DBConnStr).ExecuteScalar(sql, new {tableName, docEntry});
                return result;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return -1;
            }
        }

    }
}
