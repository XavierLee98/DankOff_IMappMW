using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class MarketingDocument
    {
        public string LogUserID { get; set; }
        public string DocObject { get; set; }
        public Int32 Series { get; set; }
        public Int32 DocNum { get; set; }
        public bool HandWritten { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime DocDueDate { get; set; }

        public int DocType { get; set; }

        public String CardCode { get; set; }
        public String CardName { get; set; }
        public String Address { get; set; }
        public String NumAtCard { get; set; }
        public String DocCurrency { get; set; }
        public Double DocRate { get; set; }
        public Double DocTotal { get; set; }
        public String Reference1 { get; set; }
        public String Reference2 { get; set; }
        public String Comments { get; set; }
        public String JournalMemo { get; set; }
        public Int32 PaymentGroupCode { get; set; }
        public DateTime DocTime { get; set; }
        public Int32 SalesPersonCode { get; set; }
        public Int32 TransportationCode { get; set; }
        public bool Confirmed { get; set; }
        public Int32 ImportFileNum { get; set; }
        public Int32 ContactPersonCode { get; set; }
        public bool ShowSCN { get; set; }
        public DateTime TaxDate { get; set; }
        public bool PartialSupply { get; set; }
        public String ShipToCode { get; set; }
        public String Indicator { get; set; }
        public String FederalTaxID { get; set; }
        public Double DiscountPercent { get; set; }
        public String PaymentReference { get; set; }
        public Int32 DocEntry { get; private set; }
        public DateTime CreationDate { get; private set; }
        public DateTime UpdateDate { get; private set; }
        public Int32 FinancialPeriod { get; private set; }
        public Int32 TransNum { get; private set; }
        public Double VatSum { get; private set; }
        public Double VatSumSys { get; private set; }
        public Double VatSumFc { get; private set; }
        public bool NetProcedure { get; private set; }
        public Double DocTotalFc { get; set; }
        public Double DocTotalSys { get; private set; }
        public Int32 Form1099 { get; set; }
        public String Box1099 { get; set; }
        public bool RevisionPo { get; set; }
        public DateTime RequriedDate { get; set; }
        public DateTime CancelDate { get; set; }
        public bool BlockDunning { get; set; }
        public bool Submitted { get; private set; }
        public Int32 Segment { get; private set; }
        public bool PickStatus { get; private set; }
        public bool Pick { get; set; }
        public String PaymentMethod { get; set; }
        public bool PaymentBlock { get; set; }
        public Int32 PaymentBlockEntry { get; set; }
        public String CentralBankIndicator { get; set; }
        public bool MaximumCashDiscount { get; set; }
        public bool Reserve { get; private set; }
        public String Project { get; set; }
        public DateTime ExemptionValidityDateFrom { get; set; }
        public DateTime ExemptionValidityDateTo { get; set; }
        public bool Rounding { get; set; }
        public String ExternalCorrectedDocNum { get; set; }
        public Int32 InternalCorrectedDocNum { get; set; }
        public Int32 NextCorrectingDocument { get; private set; }
        public bool DeferredTax { get; set; }
        public String TaxExemptionLetterNum { get; set; }
        public Double WTApplied { get; private set; }
        public Double WTAppliedFC { get; private set; }
        public bool BillOfExchangeReserved { get; private set; }
        public String AgentCode { get; set; }
        public Double WTAppliedSC { get; private set; }
        public Double TotalEqualizationTax { get; private set; }
        public Double TotalEqualizationTaxFC { get; private set; }
        public Double TotalEqualizationTaxSC { get; private set; }
        public Int32 NumberOfInstallments { get; set; }
        public bool ApplyTaxOnFirstInstallment { get; set; }
        public Double WTNonSubjectAmount { get; private set; }
        public Double WTNonSubjectAmountSC { get; private set; }
        public Double WTNonSubjectAmountFC { get; private set; }
        public Double WTExemptedAmount { get; private set; }
        public Double WTExemptedAmountSC { get; private set; }
        public Double WTExemptedAmountFC { get; private set; }
        public Double BaseAmountSC { get; private set; }
        public Double BaseAmountFC { get; private set; }
        public Double WTAmountFC { get; private set; }
        public Double WTAmountSC { get; private set; }
        public Double BaseAmount { get; private set; }
        public Double WTAmount { get; private set; }
        public DateTime VatDate { get; set; }
        public String ManualNumber { get; set; }
        public bool UseShpdGoodsAct { get; set; }
        public Int32 FolioNumber { get; set; }
        public String BPChannelCode { get; set; }
        public Int32 BPChannelContact { get; set; }
        public String DocObjectCodeEx { get; set; }
        public String Address2 { get; set; }
        public String PeriodIndicator { get; private set; }
        public String PayToCode { get; set; }
        public Int32 DocumentsOwner { get; set; }
        public String FolioPrefixString { get; set; }
        public bool IsPayToBank { get; set; }
        public String PayToBankCountry { get; set; }
        public String PayToBankCode { get; set; }
        public String PayToBankAccountNo { get; set; }
        public String PayToBankBranch { get; set; }
        public Int32 BPL_IDAssignedToInvoice { get; set; }
        public DateTime ClosingDate { get; set; }
        public Int32 SequenceSerial { get; set; }
        public Int32 SequenceCode { get; set; }
        public String SeriesString { get; set; }
        public String SubSeriesString { get; set; }
        public String SequenceModel { get; set; }
        public Double DownPayment { get; set; }
        public bool ReserveInvoice { get; set; }
        public Int32 LanguageCode { get; set; }
        public String TrackingNumber { get; set; }
        public String PickRemark { get; set; }
        public bool UseCorrectionVATGroup { get; set; }
        public Double TotalDiscount { get; private set; }
        public Double VatPercent { get; set; }
        public Double DownPaymentAmount { get; set; }
        public Double DownPaymentPercentage { get; set; }
        public Double DownPaymentAmountFC { get; set; }
        public Double DownPaymentAmountSC { get; set; }
        public String OpeningRemarks { get; set; }
        public String ClosingRemarks { get; set; }
        public Double RoundingDiffAmount { get; set; }
        public Double RoundingDiffAmountFC { get; private set; }
        public Double RoundingDiffAmountSC { get; private set; }
        public bool Cancelled { get; private set; }
        public bool InsuranceOperation347 { get; set; }
        public Double ServiceGrossProfitPercent { get; set; }
        public String ControlAccount { get; set; }
        public bool ArchiveNonremovableSalesQuotation { get; set; }
        public Int32 GTSChecker { get; set; }
        public Int32 GTSPayee { get; set; }
        public Int32 ExtraMonth { get; set; }
        public Int32 ExtraDays { get; set; }
        public Int32 CashDiscountDateOffset { get; set; }
        public bool NTSApproved { get; set; }
        public Int32 ETaxWebSite { get; set; }
        public String ETaxNumber { get; set; }
        public String NTSApprovedNumber { get; set; }
        public String SignatureInputMessage { get; private set; }
        public String SignatureDigest { get; private set; }
        public String CertificationNumber { get; private set; }
        public Int32 PrivateKeyVersion { get; private set; }
        public Int32 GroupSeries { get; set; }
        public Int32 GroupNumber { get; set; }
        public bool GroupHandWritten { get; set; }
        public bool ReopenOriginalDocument { get; set; }
        public bool ReopenManuallyClosedOrCanceledDocument { get; set; }
        public Int32 EDocSeries { get; set; }
        public String EDocNum { get; private set; }
        public Int32 EDocExportFormat { get; set; }
        public bool CreateOnlineQuotation { get; set; }
        public String POSEquipmentNumber { get; set; }
        public String POSManufacturerSerialNumber { get; set; }
        public Int32 POSCashierNumber { get; set; }
        public bool ApplyCurrentVATRatesForDownPaymentsToDraw { get; set; }
        public DateTime SpecifiedClosingDate { get; set; }
        public bool OpenForLandedCosts { get; set; }
        public Double TotalDiscountFC { get; private set; }
        public Double TotalDiscountSC { get; private set; }
        public bool RelevantToGTS { get; set; }
        public String EDocErrorCode { get; set; }
        public String EDocErrorMessage { get; set; }
        public String BPLName { get; private set; }
        public String VATRegNum { get; private set; }
        public Int32 AnnualInvoiceDeclarationReference { get; set; }
        public String Supplier { get; set; }
        public Int32 Releaser { get; set; }
        public Int32 Receiver { get; set; }
        public Int32 BlanketAgreementNumber { get; set; }
        public bool IsAlteration { get; set; }
        public DateTime AssetValueDate { get; set; }
        public String Requester { get; set; }
        public String RequesterName { get; set; }
        public Int32 RequesterBranch { get; set; }
        public Int32 RequesterDepartment { get; set; }
        public String RequesterEmail { get; set; }
        public bool SendNotification { get; set; }
        public Int32 AttachmentEntry { get; set; }
        public Int32 ReqType { get; set; }
        public bool ReuseDocumentNum { get; set; }
        public bool ReuseNotaFiscalNum { get; set; }
        public String AuthorizationCode { get; set; }
        public DateTime StartDeliveryDate { get; set; }
        public DateTime StartDeliveryTime { get; set; }
        public DateTime EndDeliveryDate { get; set; }
        public DateTime EndDeliveryTime { get; set; }
        public String VehiclePlate { get; set; }
        public String ATDocumentType { get; set; }
        public String ElecCommMessage { get; private set; }
        public bool PrintSEPADirect { get; set; }
        public String FiscalDocNum { get; set; }
    }

    public class MarketingDocumentLines
    {
        public Int32 LineNum { get; private set; }
        public String ItemCode { get; set; }
        public String ItemDescription { get; set; }
        public Double Quantity { get; set; }
        public DateTime ShipDate { get; set; }
        public Double Price { get; set; }
        public Double PriceAfterVAT { get; set; }
        public String Currency { get; set; }
        public Double Rate { get; set; }
        public Double DiscountPercent { get; set; }
        public String VendorNum { get; set; }
        public String SerialNum { get; set; }
        public String WarehouseCode { get; set; }
        public Int32 SalesPersonCode { get; set; }
        public Double CommisionPercent { get; set; }
        public String AccountCode { get; set; }
        public bool UseBaseUnits { get; set; }
        public String SupplierCatNum { get; set; }
        public String CostingCode { get; set; }
        public String ProjectCode { get; set; }
        public String BarCode { get; set; }
        public String VatGroup { get; set; }
        public Double Height1 { get; set; }
        public Int32 Hight1Unit { get; set; }
        public Double Height2 { get; set; }
        public Int32 Height2Unit { get; set; }
        public Double Lengh1 { get; set; }
        public Int32 Lengh1Unit { get; set; }
        public Double Lengh2 { get; set; }
        public Int32 Lengh2Unit { get; set; }
        public Double Weight1 { get; set; }
        public Int32 Weight1Unit { get; set; }
        public Double Weight2 { get; set; }
        public Int32 Weight2Unit { get; set; }
        public Double Factor1 { get; set; }
        public Double Factor2 { get; set; }
        public Double Factor3 { get; set; }
        public Double Factor4 { get; set; }
        public Int32 BaseType { get; set; }
        public Int32 BaseEntry { get; set; }
        public Int32 BaseLine { get; set; }
        public Double Volume { get; set; }
        public Int32 VolumeUnit { get; set; }
        public Double Width1 { get; set; }
        public Int32 Width1Unit { get; set; }
        public Double Width2 { get; set; }
        public Int32 Width2Unit { get; set; }
        public String Address { get; set; }
        public String TaxCode { get; set; }
        public bool TaxLiable { get; set; }
        public bool PickStatus { get; private set; }
        public Double PickQuantity { get; private set; }
        public Int32 PickListIdNumber { get; private set; }
        public String OriginalItem { get; private set; }
        public bool BackOrder { get; set; }
        public String FreeText { get; set; }
        public Int32 ShippingMethod { get; set; }
        public Int32 POTargetNum { get; private set; }
        public String POTargetEntry { get; private set; }
        public String POTargetRowNum { get; private set; }
        public Double CorrInvAmountToStock { get; set; }
        public Double CorrInvAmountToDiffAcct { get; set; }
        public Double AppliedTax { get; private set; }
        public Double AppliedTaxFC { get; private set; }
        public Double AppliedTaxSC { get; private set; }
        public bool WTLiable { get; set; }
        public Double EqualizationTaxPercent { get; private set; }
        public Double TotalEqualizationTaxFC { get; private set; }
        public Double TotalEqualizationTaxSC { get; private set; }
        public Double NetTaxAmount { get; set; }
        public Double NetTaxAmountFC { get; set; }
        public Double NetTaxAmountSC { get; private set; }
        public Double LineTotal { get; set; }
        public Double GrossTotal { get; set; }
        public Double TaxPercentagePerRow { get; set; }
        public Double TaxTotal { get; set; }
        public bool DeferredTax { get; set; }
        public Double TotalEqualizationTax { get; private set; }
        public String MeasureUnit { get; set; }
        public Double UnitsOfMeasurment { get; set; }
        public Double ExciseAmount { get; set; }
        public Double TaxPerUnit { get; private set; }
        public Double TotalInclTax { get; private set; }
        public String CountryOrg { get; set; }
        public String SWW { get; set; }
        public bool ConsumerSalesForecast { get; set; }
        public String CFOPCode { get; set; }
        public String CSTCode { get; set; }
        public String Usage { get; set; }
        public bool TaxOnly { get; set; }
        public Double TaxBeforeDPM { get; private set; }
        public Double TaxBeforeDPMFC { get; private set; }
        public Double TaxBeforeDPMSC { get; private set; }
        public bool DistributeExpense { get; set; }
        public String ShipToCode { get; set; }
        public Double RowTotalFC { get; set; }
        public Double RowTotalSC { get; private set; }
        public Double LastBuyInmPrice { get; private set; }
        public Double LastBuyDistributeSumFc { get; private set; }
        public Double LastBuyDistributeSumSc { get; private set; }
        public Double LastBuyDistributeSum { get; private set; }
        public Double StockDistributesumForeign { get; private set; }
        public Double StockDistributesumSystem { get; private set; }
        public Double StockDistributesum { get; private set; }
        public Double StockInmPrice { get; private set; }
        public Int32 VisualOrder { get; private set; }
        public Double BaseOpenQuantity { get; private set; }
        public Double UnitPrice { get; set; }
        public Double PackageQuantity { get; private set; }
        public String Text { get; private set; }
        public String COGSCostingCode { get; set; }
        public String COGSAccountCode { get; set; }
        public String ChangeAssemlyBoMWarehouse { get; set; }
        public String ItemDetails { get; set; }
        public DateTime ActualDeliveryDate { get; set; }
        public Int32 LocationCode { get; set; }
        public String CostingCode2 { get; set; }
        public String CostingCode3 { get; set; }
        public String CostingCode4 { get; set; }
        public String CostingCode5 { get; set; }
        public Double RemainingOpenQuantity { get; private set; }
        public Double OpenAmount { get; private set; }
        public Double OpenAmountFC { get; private set; }
        public Double OpenAmountSC { get; private set; }
        public Double GrossBuyPrice { get; set; }
        public Int32 GrossBase { get; set; }
        public Double GrossProfitTotalBasePrice { get; set; }
        public String ExLineNo { get; set; }
        public DateTime RequiredDate { get; set; }
        public Double RequiredQuantity { get; set; }
        public String COGSCostingCode2 { get; set; }
        public String COGSCostingCode3 { get; set; }
        public String COGSCostingCode4 { get; set; }
        public String COGSCostingCode5 { get; set; }
        public String CSTforIPI { get; set; }
        public String CSTforPIS { get; set; }
        public String CSTforCOFINS { get; set; }
        public String CreditOriginCode { get; set; }
        public bool WithoutInventoryMovement { get; set; }
        public Int32 AgreementNo { get; set; }
        public Int32 AgreementRowNumber { get; set; }
        public String ShipToDescription { get; set; }
        public Int32 ActualBaseEntry { get; set; }
        public Int32 ActualBaseLine { get; set; }
        public Int32 DocEntry { get; private set; }
        public Double Surpluses { get; set; }
        public Double DefectAndBreakup { get; set; }
        public Double Shortages { get; set; }
        public bool ConsiderQuantity { get; set; }
        public bool PartialRetirement { get; set; }
        public Double RetirementQuantity { get; set; }
        public Double RetirementAPC { get; set; }
        public bool EnableReturnCost { get; set; }
        public Double ReturnCost { get; set; }
        public Int32 UoMEntry { get; set; }
        public String UoMCode { get; private set; }
        public Double InventoryQuantity { get; set; }
        public Double RemainingOpenInventoryQuantity { get; private set; }
        public String LineVendor { get; set; }
        public Int32 Incoterms { get; set; }
        public Int32 TransportMode { get; set; }
    }

    public class MarketingDocumentLinesSerials
    {
        public String ManufacturerSerialNumber { get; set; }
        public String InternalSerialNumber { get; set; }
        public DateTime ExpiryDate { get; set; }
        public DateTime ManufactureDate { get; set; }
        public DateTime ReceptionDate { get; set; }
        public DateTime WarrantyStart { get; set; }
        public DateTime WarrantyEnd { get; set; }
        public String Location { get; set; }
        public String Notes { get; set; }
        public String BatchID { get; set; }
        public Int32 SystemSerialNumber { get; set; }
        public Int32 BaseLineNumber { get; set; }
        public Double Quantity { get; set; }
        public Int32 TrackingNote { get; set; }
        public Int32 TrackingNoteLine { get; set; }
    }
    public class MarketingDocumentLinesBatch
    {
        public String BatchNumber { get; set; }
        public String ManufacturerSerialNumber { get; set; }
        public String InternalSerialNumber { get; set; }
        public DateTime ExpiryDate { get; set; }
        public DateTime ManufacturingDate { get; set; }
        public DateTime AddmisionDate { get; set; }
        public String Location { get; set; }
        public String Notes { get; set; }
        public Double Quantity { get; set; }
        public Int32 BaseLineNumber { get; set; }
        public Int32 TrackingNote { get; set; }
        public Int32 TrackingNoteLine { get; set; }
    }

    public class MarketingDocumentDownPayment
    {
        public Int32 Count { get; private set; }
        public Int32 DocEntry { get; set; }
        public DateTime PostingDate { get; private set; }
        public DateTime DueDate { get; private set; }
        public String Name { get; private set; }
        public String Details { get; private set; }
        public Double AmountToDraw { get; set; }
        public Double AmountToDrawFC { get; set; }
        public Double AmountToDrawSC { get; set; }
        public Int32 RowNum { get; private set; }
        public Int32 DocNumber { get; private set; }
        public Double Tax { get; private set; }
        public Double TaxFC { get; private set; }
        public Double TaxSC { get; private set; }
        public Double GrossAmountToDraw { get; set; }
        public Double GrossAmountToDrawFC { get; set; }
        public Double GrossAmountToDrawSC { get; set; }
        public bool IsGrossLine { get; private set; }
        public Int32 DocInternalID { get; private set; }
    }
}
