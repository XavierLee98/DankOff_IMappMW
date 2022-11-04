using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class InventoryTransfer
    {
        public string LogUserID { get; set; }
        public Int32 Series { get; set; }
        public bool Printed { get; private set; }
        public DateTime DocDate { get; set; }
        public String CardCode { get; set; }
        public String CardName { get; set; }
        public String Address { get; set; }
        public String Reference1 { get; set; }
        public String Reference2 { get; set; }
        public String Comments { get; set; }
        public String JournalMemo { get; set; }
        public Int32 PriceList { get; set; }
        public Int32 SalesPersonCode { get; set; }
        public String FromWarehouse { get; set; }
        public Int32 DocEntry { get; private set; }
        public DateTime CreationDate { get; private set; }
        public DateTime UpdateDate { get; private set; }
        public Int32 FinancialPeriod { get; private set; }
        public Int32 TransNum { get; private set; }
        public Int32 DocNum { get; private set; }
        public DateTime TaxDate { get; set; }
        public Int32 ContactPerson { get; set; }
        public String FolioPrefixString { get; set; }
        public Int32 FolioNumber { get; set; }
        public DateTime DueDate { get; set; }
        public Int32 BPLID { get; private set; }
        public String BPLName { get; private set; }
        public String VATRegNum { get; private set; }
        public String ToWarehouse { get; set; }
        public String AuthorizationCode { get; set; }
        public DateTime StartDeliveryDate { get; set; }
        public DateTime StartDeliveryTime { get; set; }
        public DateTime EndDeliveryDate { get; set; }
        public DateTime EndDeliveryTime { get; set; }
        public String VehiclePlate { get; set; }
        public String ATDocumentType { get; set; }
        public Int32 EDocExportFormat { get; set; }
        public String ElecCommMessage { get; private set; }
    }
    public class InventoryTransferLines
    {
        public Int32 LineNum { get; private set; }
        public String ItemCode { get; set; }
        public String ItemDescription { get; set; }
        public Double Quantity { get; set; }
        public Double Price { get; set; }
        public String Currency { get; set; }
        public Double Rate { get; set; }
        public Double DiscountPercent { get; set; }
        public String VendorNum { get; set; }
        public String SerialNumber { get; set; }
        public String WarehouseCode { get; set; }
        public String ProjectCode { get; set; }
        public Double Factor { get; set; }
        public Double Factor2 { get; set; }
        public Double Factor3 { get; set; }
        public Double Factor4 { get; set; }
        public String DistributionRule { get; set; }
        public bool UseBaseUnits { get; set; }
        public String MeasureUnit { get; set; }
        public Double UnitsOfMeasurment { get; set; }
        public String DistributionRule2 { get; set; }
        public String DistributionRule3 { get; set; }
        public String DistributionRule4 { get; set; }
        public String DistributionRule5 { get; set; }
        public Int32 BaseLine { get; set; }
        public Int32 BaseEntry { get; set; }
        public Int32 DocEntry { get; private set; }
        public Double UnitPrice { get; set; }
        public String FromWarehouseCode { get; set; }
        public Int32 UoMEntry { get; set; }
        public String UoMCode { get; private set; }
        public Double InventoryQuantity { get; set; }
        public Double RemainingOpenQuantity { get; private set; }
        public Double RemainingOpenInventoryQuantity { get; private set; }
    }
}
