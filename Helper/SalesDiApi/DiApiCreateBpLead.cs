using Dapper;
using Microsoft.Data.SqlClient;
using IMAppSapMidware_NetCore.Models.BpLead;
using IMAppSapMidware_NetCore.Models.Share;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IMAppSapMidware_NetCore.Helper.DiApi
{
    public class DiApiCreateBpLead
    {   
        public ErpPropertyHelper ErpProperty { get; set; } = null;
        public string LastSAPMsg { get; set; } = string.Empty;              
        public string Midware_DBConnStr { get; set; } = string.Empty;
        public string Erp_DBConnStr { get; set; } = string.Empty;
        public string Guid { get; set; } = string.Empty;
        public string DocType { get; set; } = string.Empty;
        public int AttachedFileCnt { get; set; } = -1;
        public BpNewLead BpLead { get; set; } = null;
        public List<BpNewLeadShipBillAddr> BpAddresses { get; set; } = null;
        public List<BpNewLeadContactPerson> BpContacts { get; set; } = null;
        private Company sapCompany { get; set; }


        void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }

        public bool CreateBpLeadDocument()
        {
            string modName = "[CreateBpLeadDocument][" + DocType + "]";
            bool retResult = false;
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

                var bp = (BusinessPartners)sapCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                var curCardCode = FormatValue(BpLead.Cardcode, 15);
                bp.CardCode = curCardCode;
                bp.CardName = FormatValue(BpLead.CardName, 100);
                bp.CardForeignName = FormatValue(BpLead.CardFName, 100);
                bp.CardType = BoCardTypes.cLid; // always save as lead

                var curCode = GetCurrencyCode(BpLead.Currency); // return 3 char currency string
                if (!string.IsNullOrWhiteSpace(curCode))
                {
                    bp.Currency = curCode;
                }

                bp.Phone1 = FormatValue(BpLead.Phone1, 20);
                bp.Phone2 = FormatValue(BpLead.Phone2, 20);
                bp.Cellular = FormatValue(BpLead.Cellular, 20);
                bp.Fax = FormatValue(BpLead.Fax, 20);
                bp.EmailAddress = FormatValue(BpLead.E_Mail, 100);
                bp.Website = FormatValue(BpLead.IntrntSite, 100);

                int shipType = GetShippingTypeCode(BpLead.ShippingType);
                if (shipType >= 0)
                {
                    bp.ShippingType = shipType;
                }

                bp.CompanyRegistrationNumber = FormatValue(BpLead.RegNum, 32);


                var slpCode = GetSalesPersonCode(BpLead.SlpCode);
                if (slpCode >= 0)
                {
                    bp.SalesPersonCode = slpCode;
                }

                // add in the ship and bill address by loop
                if (BpAddresses != null)
                {
                    BpAddresses.ForEach(address =>
                    {
                        bp.Addresses.AddressName = FormatValue(address.Address, 50);
                        bp.Addresses.Street = FormatValue(address.Street, 100);
                        bp.Addresses.Block = FormatValue(address.Block, 100);
                        bp.Addresses.ZipCode = FormatValue(address.ZipCode, 20);
                        //bp.Addresses.State = ut.SafeGetStr(addrRow["state"]); // tempporary hide
                        bp.Addresses.BuildingFloorRoom = address.Building;
                        bp.Addresses.StreetNo = address.StreetNo;
                        bp.Addresses.City = FormatValue(address.City, 100);
                        var country = GetTwoCharCountryCode(FormatValue(address.Country, 100));
                        bp.Addresses.Country = country;
                        bp.Addresses.AddressType = (address.AdresType.Equals("B")) // ut.SafeGetStr(addrRow["AdresType"]).Equals("B"))
                            ? BoAddressType.bo_BillTo
                            : BoAddressType.bo_ShipTo;
                        bp.Addresses.AddressName2 = FormatValue(address.Address2, 50);//  ut.SafeGetStr(addrRow["Address2"]), 50);
                        bp.Addresses.AddressName3 = FormatValue(address.Address3, 50);//  ut.SafeGetStr(addrRow["Address3"]), 50);
                        bp.Addresses.Add();

                    
                    });
                }

                // add in the ship and contact person by loop
                if (BpContacts != null)
                {
                    BpContacts.ForEach(contact =>
                   {
                       bp.ContactEmployees.Name = FormatValue(contact.Name, 50); //  ut.SafeGetStr(contactRow["Name"]), 50);
                       bp.ContactEmployees.Position = FormatValue(contact.Position, 90);//  ut.SafeGetStr(contactRow["Position"]), 90);
                       bp.ContactEmployees.Address = FormatValue(contact.Address, 100); //  ut.SafeGetStr(contactRow["Address"]), 100);
                       bp.ContactEmployees.Phone1 = FormatValue(contact.Tel1, 20); // // ut.SafeGetStr(contactRow["Tel1"]), 20);
                       bp.ContactEmployees.Phone2 = FormatValue(contact.Tel2, 20); // ut.SafeGetStr(contactRow["Tel2"]), 20);
                       bp.ContactEmployees.MobilePhone = FormatValue(contact.Cellolar, 50);  //ut.SafeGetStr(contactRow["Cellolar"]), 50);
                       bp.ContactEmployees.Fax = FormatValue(contact.Fax, 20); // ut.SafeGetStr(contactRow["Fax"]), 20);
                       bp.ContactEmployees.E_Mail = FormatValue(contact.E_MailL, 100); // ut.SafeGetStr(contactRow["E_MailL"]), 100);
                       bp.ContactEmployees.Pager = FormatValue(contact.Pager, 30); // ut.SafeGetStr(contactRow["Pager"]), 30);

                       var countryOfBirth = GetTwoCharCountryCode(FormatValue(contact.BirthPlace, 100));
                       bp.ContactEmployees.PlaceOfBirth = countryOfBirth;
                       bp.ContactEmployees.DateOfBirth = contact.BirthDate; // ut.SafeGetDateTime(contactRow["BirthDate"]);
                       bp.ContactEmployees.Gender = GetGender(contact.Gender);//  ut.SafeGetStr(contactRow["Gender"]));
                       bp.ContactEmployees.Profession = FormatValue(contact.Profession, 50); ///ut.SafeGetStr(contactRow["Profession"]), 50);
                       bp.ContactEmployees.Title = FormatValue(contact.Title, 10);// ut.SafeGetStr(contactRow["Title"]), 10);
                       bp.ContactEmployees.CityOfBirth = FormatValue(contact.BirthCity, 100); // ut.SafeGetStr(contactRow["BirthCity"]), 100);
                       bp.ContactEmployees.Add();
                   });
                }

                if (bp.Add() == 0) // success
                {
                    if (AttachedFileCnt > 0)
                    {
                        AddFileAttachment(BpLead.Cardcode);
                    }
                    
                    Log($"{modName} Created, Bp: {BpLead.Cardcode}, {BpLead.CardName} , {BpLead.Guid}");
                    return true;
                }

                // else                
                Log($"{modName}\n{sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
                return false;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");                                
                return retResult;
            }
        }

        void AddFileAttachment(string cardCode)
        {
            try
            {
                var bp = (BusinessPartners)sapCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                if (!bp.GetByKey(cardCode)) return; // no bp object found

                var files = GetAttachmentFiles(this.Guid);
                if (files == null) return;
                if (files.Count == 0) return;

                var attachment = (Attachments2)sapCompany.GetBusinessObject(BoObjectTypes.oAttachments2);
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
                    Log($"BpLead Attachments {sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
                    return;
                }

                if (attachment.Add() == 0)
                {
                    int iAttEntry = int.Parse(sapCompany.GetNewObjectKey());
                    bp.AttachmentEntry = iAttEntry;
                    bp.Update();
                    return;
                }

                // else                    
                Log($"BpLead Attachments {sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
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

        string GetCurrencyCode(string currency)
        {
            try
            {
                var query = "SELECT TOP 1 currCode FROM OCRN " +
                    "WHERE currCode=@currCode";

                var result = string.Empty;
                using (var conn = new SqlConnection(this.Erp_DBConnStr))
                {
                    result =  conn.ExecuteScalar<string>(query, new { currCode = currency }) ;
                }
               
                // check result, if not foun then query default company currency 
                if (string.IsNullOrWhiteSpace(result))
                {
                    result = GetDefaultCompanyCurrency();
                }

                return result;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return GetDefaultCompanyCurrency();
            }            
        }

        string GetDefaultCompanyCurrency()
        {            
            try
            {                
                return new SqlConnection(this.Erp_DBConnStr).ExecuteScalar<string>("SELECT TOP 1 MainCurncy FROM OADM");
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return string.Empty;
            }
        }


        int GetShippingTypeCode(string trnspName)
        {
            try
            {
                return new SqlConnection(this.Erp_DBConnStr).ExecuteScalar<int>(
                    @"SELECT TrnspCode FROM OSHP WHERE TrnspName Like '%@TrnspName%'", new { TrnspName= trnspName });
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return -1;
            }
        }

        int GetSalesPersonCode(string salesPersonName)
        {            
            try
            {
                if (string.IsNullOrWhiteSpace(salesPersonName)) return -1;
                if (salesPersonName.ToLower().Equals("-No Sales Employee-".ToLower()))
                {
                    return -1;
                }   

                return new SqlConnection(this.Erp_DBConnStr).ExecuteScalar<int>(
                   @"SELECT SlpCode FROM OSLP WHERE SlpName = @SlpName", new { SlpName = salesPersonName });
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return -1;
            }
        }

        string GetTwoCharCountryCode(string countryName)
        {
            try
            {
                var result = new SqlConnection(this.Erp_DBConnStr).ExecuteScalar<string>(
                         @"SELECT TOP 1 Code FROM OCRY WHERE Name Like '%@country%'", new { country = countryName });

                if (string.IsNullOrWhiteSpace(result)) result = "MY";  /// default the country to malaysia                
                return result;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return "MY";
            }
        }

        List<FileUpload> GetAttachmentFiles(string guid) // query from midware
        {
            try
            {
                var sqlQuery = @"SELECT * 
                                 FROM zmwFileUpload 
                                 WHERE HeaderGuid=@guid";

                return new SqlConnection(Midware_DBConnStr).Query<FileUpload>(sqlQuery, new { guid }).ToList();
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return null;
            }
        }

        BoGenderTypes GetGender(string gender)
        {
            switch (gender)
            {
                case "M": return BoGenderTypes.gt_Male;
                case "F": return BoGenderTypes.gt_Female;
                default: return BoGenderTypes.gt_Undefined;
            }
        }
    }
}
