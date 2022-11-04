using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class BusinessPartnerMaster
    {
        public string LogUserID { get; set; }
        public String CardCode { get; set; }
        public String CardName { get; set; }
        public Int32 GroupCode { get; set; }
        public String Address { get; set; }
        public String ZipCode { get; set; }
        public String MailAddress { get; set; }
        public String MailZipCode { get; set; }
        public String Phone1 { get; set; }
        public String Phone2 { get; set; }
        public String Fax { get; set; }
        public String ContactPerson { get; set; }
        public String Notes { get; set; }
        public Int32 PayTermsGrpCode { get; set; }
        public Double CreditLimit { get; set; }
        public Double MaxCommitment { get; set; }
        public Double DiscountPercent { get; set; }
        public String FederalTaxID { get; set; }
        public bool DeductibleAtSource { get; set; }
        public Double DeductionPercent { get; set; }
        public DateTime DeductionValidUntil { get; set; }
        public Int32 PriceListNum { get; set; }
        public Double IntrestRatePercent { get; set; }
        public Double CommissionPercent { get; set; }
        public Int32 CommissionGroupCode { get; set; }
        public String FreeText { get; set; }
        public Int32 SalesPersonCode { get; set; }
        public String Currency { get; set; }
        public String RateDiffAccount { get; set; }
        public String Cellular { get; set; }
        public Int32 AvarageLate { get; set; }
        public String City { get; set; }
        public String County { get; set; }
        public String Country { get; set; }
        public String MailCity { get; set; }
        public String MailCounty { get; set; }
        public String MailCountry { get; set; }
        public String EmailAddress { get; set; }
        public String Picture { get; set; }
        public String DefaultAccount { get; set; }
        public String DefaultBranch { get; set; }
        public String DefaultBankCode { get; set; }
        public String AdditionalID { get; set; }
        public String Pager { get; set; }
        public String FatherCard { get; set; }
        public String CardForeignName { get; set; }
        public bool Properties { get; set; }
        public String DeductionOffice { get; set; }
        public String ExportCode { get; set; }
        public Double MinIntrest { get; set; }
        public Double CurrentAccountBalance { get; private set; }
        public Double OpenDeliveryNotesBalance { get; private set; }
        public Double OpenOrdersBalance { get; private set; }
        public String VatGroup { get; set; }
        public Int32 ShippingType { get; set; }
        public String Password { get; set; }
        public String Indicator { get; set; }
        public String IBAN { get; set; }
        public Int32 CreditCardCode { get; set; }
        public String CreditCardNum { get; set; }
        public DateTime CreditCardExpiration { get; set; }
        public String DebitorAccount { get; set; }
        public Int32 OpenOpportunities { get; private set; }
        public bool Valid { get; set; }
        public DateTime ValidFrom { get; set; }
        public DateTime ValidTo { get; set; }
        public String ValidRemarks { get; set; }
        public bool Frozen { get; set; }
        public DateTime FrozenFrom { get; set; }
        public DateTime FrozenTo { get; set; }
        public String FrozenRemarks { get; set; }
        public String Block { get; set; }
        public String BillToState { get; set; }
        public String ExemptNum { get; set; }
        public Int32 Priority { get; set; }
        public Int32 FormCode1099 { get; set; }
        public String Box1099 { get; set; }
        public String PeymentMethodCode { get; set; }
        public bool BackOrder { get; set; }
        public bool PartialDelivery { get; set; }
        public bool BlockDunning { get; set; }
        public String BankCountry { get; set; }
        public String HouseBank { get; set; }
        public String HouseBankCountry { get; set; }
        public String HouseBankAccount { get; set; }
        public String ShipToDefault { get; set; }
        public Int32 DunningLevel { get; private set; }
        public DateTime DunningDate { get; private set; }
        public bool CollectionAuthorization { get; set; }
        public String DME { get; set; }
        public String InstructionKey { get; set; }
        public bool SinglePayment { get; set; }
        public String ISRBillerID { get; set; }
        public bool PaymentBlock { get; set; }
        public String ReferenceDetails { get; set; }
        public String HouseBankBranch { get; set; }
        public String OwnerIDNumber { get; set; }
        public Int32 PaymentBlockDescription { get; set; }
        public String TaxExemptionLetterNum { get; set; }
        public Double MaxAmountOfExemption { get; set; }
        public DateTime ExemptionValidityDateFrom { get; set; }
        public DateTime ExemptionValidityDateTo { get; set; }
        public String LinkedBusinessPartner { get; set; }
        public bool Equalization { get; set; }
        public bool SubjectToWithholdingTax { get; set; }
        public String CertificateNumber { get; set; }
        public DateTime ExpirationDate { get; set; }
        public String NationalInsuranceNum { get; set; }
        public bool AccrualCriteria { get; set; }
        public String WTCode { get; set; }
        public bool DeferredTax { get; set; }
        public String BillToBuildingFloorRoom { get; set; }
        public String DownPaymentClearAct { get; set; }
        public String ChannelBP { get; set; }
        public Int32 DefaultTechnician { get; set; }
        public String BilltoDefault { get; set; }
        public String CustomerBillofExchangDisc { get; set; }
        public Int32 Territory { get; set; }
        public String ShipToBuildingFloorRoom { get; set; }
        public String CustomerBillofExchangPres { get; set; }
        public String ProjectCode { get; set; }
        public String VatGroupLatinAmerica { get; set; }
        public String DunningTerm { get; set; }
        public String Website { get; set; }
        public String OtherReceivablePayable { get; set; }
        public Int32 ClosingDateProcedureNumber { get; set; }
        public String Profession { get; set; }
        public String BillofExchangeonCollection { get; set; }
        public Int32 LanguageCode { get; set; }
        public String UnpaidBillofExchange { get; set; }
        public Int32 WithholdingTaxDeductionGroup { get; set; }
        public String BankChargesAllocationCode { get; set; }
        public String CompanyRegistrationNumber { get; set; }
        public String VerificationNumber { get; set; }
        public bool InsuranceOperation347 { get; set; }
        public bool ThresholdOverlook { get; set; }
        public bool SurchargeOverlook { get; set; }
        public String DownPaymentInterimAccount { get; set; }
        public bool HierarchicalDeduction { get; set; }
        public bool WithholdingTaxCertified { get; set; }
        public bool BookkeepingCertified { get; set; }
        public String PlanningGroup { get; set; }
        public bool Affiliate { get; set; }
        public Int32 Industry { get; set; }
        public String VatIDNum { get; set; }
        public Int32 DatevAccount { get; set; }
        public bool DatevFirstDataEntry { get; set; }
        public String GTSRegNo { get; set; }
        public String GTSBankAccountNo { get; set; }
        public String GTSBillingAddrTel { get; set; }
        public Int32 ETaxWebSite { get; set; }
        public String HouseBankIBAN { get; private set; }
        public Int32 Series { get; set; }
        public String InterestAccount { get; set; }
        public String FeeAccount { get; set; }
        public Int32 CampaignNumber { get; set; }
        public String VATRegistrationNumber { get; set; }
        public String RepresentativeName { get; set; }
        public String IndustryType { get; set; }
        public String BusinessType { get; set; }
        public String AliasName { get; set; }
        public Int32 DefaultBlanketAgreementNumber { get; set; }
        public bool NoDiscounts { get; set; }
        public String GlobalLocationNumber { get; set; }
        public String EDISenderID { get; set; }
        public String EDIRecipientID { get; set; }
        public String UnifiedFederalTaxID { get; set; }
        public DateTime RelationshipDateFrom { get; set; }
        public DateTime RelationshipDateTill { get; set; }
        public String RelationshipCode { get; set; }
        public Int32 AttachmentEntry { get; set; }
        public int AutomaticPosting { get; set; }
        private int _VatLiable = 1;
        public int VatLiable { get { return _VatLiable; } set { _VatLiable = value; } }

        public int CardType { get; set; }
        public int CompanyPrivate { get; set; }
        public int DiscountBaseObject { get; set; }
        public int DiscountRelations { get; set; }
        public int EffectiveDiscount { get; set; }
        public int FatherType { get; set; }
    }

    public class BusinessPartnerMasterAddress
    {
        public String AddressName { get; set; }
        public String Street { get; set; }
        public String Block { get; set; }
        public String ZipCode { get; set; }
        public String City { get; set; }
        public String County { get; set; }
        public String Country { get; set; }
        public String State { get; set; }
        public String FederalTaxID { get; set; }
        public String TaxCode { get; set; }
        public String BuildingFloorRoom { get; set; }
        public String AddressName2 { get; set; }
        public String AddressName3 { get; set; }
        public String TypeOfAddress { get; set; }
        public String StreetNo { get; set; }
        public String BPCode { get; private set; }
        public Int32 RowNum { get; private set; }
        public String GlobalLocationNumber { get; set; }
        public int AddressType { get; set; }

    }

    public class BusinesspartnerMasterContact
    {
        public String Position { get; set; }
        public String Address { get; set; }
        public String Phone1 { get; set; }
        public String Phone2 { get; set; }
        public String MobilePhone { get; set; }
        public String Fax { get; set; }
        public String E_Mail { get; set; }
        public String Pager { get; set; }
        public String Remarks1 { get; set; }
        public String Remarks2 { get; set; }
        public String Password { get; set; }
        public String Name { get; set; }
        public Int32 InternalCode { get; private set; }
        public String PlaceOfBirth { get; set; }
        public DateTime DateOfBirth { get; set; }
        public String Profession { get; set; }
        public String CardCode { get; private set; }
        public String Title { get; set; }
        public String CityOfBirth { get; set; }
        public bool Active { get; set; }
        public String FirstName { get; set; }
        public String MiddleName { get; set; }
        public String LastName { get; set; }
        private int _Gender = 2;
        public int Gender { get { return _Gender; } set { _Gender = value; } }
    }
}
