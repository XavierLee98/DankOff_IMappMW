using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class ItemMaster
    {
        public string LogUserID { get; set; }
        public String ItemCode { get; set; }
        public String ItemName { get; set; }
        public String ForeignName { get; set; }
        public Int32 ItemsGroupCode { get; set; }
        public Int32 CustomsGroupCode { get; set; }
        public String SalesVATGroup { get; set; }
        public String BarCode { get; set; }
        public bool VatLiable { get; set; }

        private bool _PurchaseItem = true;
        public bool PurchaseItem { get { return _PurchaseItem; } set { _PurchaseItem = value; } }

        private bool _SalesItem = true;
        public bool SalesItem { get { return _SalesItem; } set { _SalesItem = value; } }

        private bool _InventoryItem = true;
        public bool InventoryItem { get { return _InventoryItem; } set { _InventoryItem = value; } }
        private int _GLMethod = 1;
        public int GLMethod { get { return _GLMethod; } set { _GLMethod = value; } }
        public String IncomeAccount { get; set; }
        public String Mainsupplier { get; set; }
        public String SupplierCatalogNo { get; set; }
        public String PurchaseUnit { get; set; }
        public Double PurchaseItemsPerUnit { get; set; }
        public Double DesiredInventory { get; set; }
        public Double MinInventory { get; set; }
        public String SalesUnit { get; set; }
        public Double SalesItemsPerUnit { get; set; }
        public String Picture { get; set; }
        public String User_Text { get; set; }
        public String SerialNum { get; set; }
        public Double CommissionPercent { get; set; }
        public Double CommissionSum { get; set; }
        public Int32 CommissionGroup { get; set; }
        public bool AssetItem { get; set; }
        public Double SalesUnitHeight { get; set; }
        public Double SalesUnitVolume { get; set; }
        public String DataExportCode { get; set; }
        public bool Properties { get; set; }
        public Int32 Manufacturer { get; set; }
        public Double QuantityOnStock { get; private set; }
        public Double QuantityOrderedFromVendors { get; private set; }
        public Double QuantityOrderedByCustomers { get; private set; }
        public bool ManageSerialNumbers { get; set; }
        public bool ManageBatchNumbers { get; set; }
        public bool Valid { get; set; }
        public DateTime ValidFrom { get; set; }
        public DateTime ValidTo { get; set; }
        public String ValidRemarks { get; set; }
        public bool Frozen { get; set; }
        public DateTime FrozenFrom { get; set; }
        public DateTime FrozenTo { get; set; }
        public String FrozenRemarks { get; set; }
        public String SalesPackagingUnit { get; set; }
        public Double SalesQtyPerPackUnit { get; set; }
        public Double SalesUnitLength { get; set; }
        public Double SalesUnitWidth { get; set; }
        public Int32 SalesVolumeUnit { get; set; }
        public Double SalesUnitWeight { get; set; }
        public String PurchasePackagingUnit { get; set; }
        public Double PurchaseQtyPerPackUnit { get; set; }
        public Double PurchaseUnitLength { get; set; }
        public Double PurchaseUnitWidth { get; set; }
        public Double PurchaseUnitHeight { get; set; }
        public Double PurchaseUnitVolume { get; set; }
        public Int32 PurchaseVolumeUnit { get; set; }
        public Double PurchaseUnitWeight { get; set; }
        public String PurchaseVATGroup { get; set; }
        public Double SalesFactor1 { get; set; }
        public Double SalesFactor2 { get; set; }
        public Double SalesFactor3 { get; set; }
        public Double SalesFactor4 { get; set; }
        public Double PurchaseFactor1 { get; set; }
        public Double PurchaseFactor2 { get; set; }
        public Double PurchaseFactor3 { get; set; }
        public Double PurchaseFactor4 { get; set; }
        public Double MovingAveragePrice { get; private set; }
        public Int32 SalesLengthUnit { get; set; }
        public Int32 SalesWidthUnit { get; set; }
        public Int32 SalesHeightUnit { get; set; }
        public Int32 SalesWeightUnit { get; set; }
        public Int32 PurchaseWidthUnit { get; set; }
        public Int32 PurchaseLengthUnit { get; set; }
        public Int32 PurchaseHeightUnit { get; set; }
        public Int32 PurchaseWeightUnit { get; set; }
        public String ForeignRevenuesAccount { get; set; }
        public String ECRevenuesAccount { get; set; }
        public String ForeignExpensesAccount { get; set; }
        public String ECExpensesAccount { get; set; }
        public Double AvgStdPrice { get; set; }
        public String DefaultWarehouse { get; set; }
        public Int32 ShipType { get; set; }
        public Double MaxInventory { get; set; }

        private bool _ManageStockByWarehouse = false;
        public bool ManageStockByWarehouse { get { return _ManageStockByWarehouse; } set { _ManageStockByWarehouse = value; } }
        public Int32 PurchaseHeightUnit1 { get; set; }
        public Double PurchaseUnitHeight1 { get; set; }
        public Int32 PurchaseLengthUnit1 { get; set; }
        public Double PurchaseUnitLength1 { get; set; }
        public Int32 PurchaseWeightUnit1 { get; set; }
        public Double PurchaseUnitWeight1 { get; set; }
        public Int32 PurchaseWidthUnit1 { get; set; }
        public Double PurchaseUnitWidth1 { get; set; }
        public Int32 SalesHeightUnit1 { get; set; }
        public Double SalesUnitHeight1 { get; set; }
        public Int32 SalesLengthUnit1 { get; set; }
        public Double SalesUnitLength1 { get; set; }
        public Int32 SalesWeightUnit1 { get; set; }
        public Double SalesUnitWeight1 { get; set; }
        public Int32 SalesWidthUnit1 { get; set; }
        public Double SalesUnitWidth1 { get; set; }
        public bool ForceSelectionOfSerialNumber { get; set; }
        public bool ManageSerialNumbersOnReleaseOnly { get; set; }
        public bool WTLiable { get; set; }
        public String BaseUnitName { get; set; }
        public String ItemCountryOrg { get; set; }
        public String SWW { get; set; }
        public String ApTaxCode { get; set; }
        public String ArTaxCode { get; set; }
        public String WarrantyTemplate { get; set; }
        public bool IndirectTax { get; set; }
        public bool IsPhantom { get; set; }
        public String InventoryUOM { get; set; }
        public String OrderIntervals { get; set; }
        public Double OrderMultiple { get; set; }
        public Int32 LeadTime { get; set; }
        public Double MinOrderQuantity { get; set; }
        public Int32 OutgoingServiceCode { get; set; }
        public Int32 IncomingServiceCode { get; set; }
        public Int32 ServiceGroup { get; set; }
        public Int32 NCMCode { get; set; }
        public Int32 MaterialGroup { get; set; }
        public Int32 ProductSource { get; set; }
        public bool AutoCreateSerialNumbersOnRelease { get; set; }
        public Int32 DNFEntry { get; set; }
        public String GTSItemSpec { get; set; }
        public String GTSItemTaxCategory { get; set; }
        public Int32 FuelID { get; set; }
        public String BeverageTableCode { get; set; }
        public String BeverageGroupCode { get; set; }
        public Int32 BeverageCommercialBrandCode { get; set; }
        public Int32 Series { get; set; }
        public Int32 ToleranceDays { get; set; }
        public bool NoDiscounts { get; set; }
        public String AssetClass { get; set; }
        public String AssetGroup { get; set; }
        public String InventoryNumber { get; set; }
        public Int32 Technician { get; set; }
        public Int32 Employee { get; set; }
        public Int32 Location { get; set; }
        public bool StatisticalAsset { get; set; }
        public bool Cession { get; set; }
        public bool DeactivateAfterUsefulLife { get; set; }
        public DateTime CapitalizationDate { get; set; }
        public bool ManageByQuantity { get; private set; }
        public Int32 UoMGroupEntry { get; set; }
        public Int32 InventoryUoMEntry { get; set; }
        public Int32 DefaultSalesUoMEntry { get; set; }
        public Int32 DefaultPurchasingUoMEntry { get; set; }
        public String DepreciationGroup { get; set; }
        public String AssetSerialNumber { get; set; }
        public String DefaultCountingUnit { get; private set; }
        public Double CountingItemsPerUnit { get; private set; }
        public Int32 DefaultCountingUoMEntry { get; set; }

        public int CostAccountingMethod { get; set; }
        public bool Excisable { get; set; }
        public Int32 ChapterID { get; set; }
        public String ScsCode { get; set; }

        public int IssueMethod { get; set; }
        public int IssuePrimarilyBy { get; set; }
        public int ItemType { get; set; }

        private int _PlanningSystem = 1;
        public int PlanningSystem { get { return _PlanningSystem; } set { _PlanningSystem = value; } }
        public int ProcurementMethod { get; set; }
        public int SRIAndBatchManageMethod { get; set; }
    }

    public class ItemMasterPriceList
    {
        public Int32 PriceList { get; private set; }
        public String PriceListName { get; private set; }
        public Double Price { get; set; }
        public String Currency { get; set; }
        public Double AdditionalPrice1 { get; set; }
        public String AdditionalCurrency1 { get; set; }
        public Double AdditionalPrice2 { get; set; }
        public String AdditionalCurrency2 { get; set; }
    }

    public class ItemMasterWarehouse
    {
        public Double MaximalStock { get; set; }
        public Double MinimalOrder { get; set; }
        public Double StandardAveragePrice { get; set; }
        public bool Locked { get; set; }
        public String InventoryAccount { get; set; }
        public String CostAccount { get; set; }
        public String TransferAccount { get; set; }
        public String RevenuesAccount { get; set; }
        public String VarienceAccount { get; set; }
        public String DecreasingAccount { get; set; }
        public String IncreasingAccount { get; set; }
        public String ReturningAccount { get; set; }
        public String ExpensesAccount { get; set; }
        public String EURevenuesAccount { get; set; }
        public String EUExpensesAccount { get; set; }
        public String ForeignRevenueAcc { get; set; }
        public String ForeignExpensAcc { get; set; }
        public String ExemptIncomeAcc { get; set; }
        public String PriceDifferenceAcc { get; set; }
        public Double MinimalStock { get; set; }
        public String WarehouseCode { get; set; }
        public Double InStock { get; private set; }
        public Double Committed { get; private set; }
        public Double Ordered { get; private set; }
        public Double CountedQuantity { get; private set; }
        public bool WasCounted { get; private set; }
        public Double Counted { get; private set; }
        public String WipVarianceAccount { get; set; }
        public String ExpenseClearingAct { get; set; }
        public String PurchaseCreditAcc { get; set; }
        public String EUPurchaseCreditAcc { get; set; }
        public String ForeignPurchaseCreditAcc { get; set; }
        public String SalesCreditAcc { get; set; }
        public String SalesCreditEUAcc { get; set; }
        public String ExemptedCredits { get; set; }
        public String SalesCreditForeignAcc { get; set; }
        public String ExpenseOffsettingAccount { get; set; }
        public String WipAccount { get; set; }
        public String GoodsClearingAcct { get; set; }
        public String ExchangeRateDifferencesAcct { get; set; }
        public String NegativeInventoryAdjustmentAccount { get; set; }
        public String CostInflationAccount { get; set; }
        public String CostInflationOffsetAccount { get; set; }
        public String GLDecreaseAcct { get; set; }
        public String GLIncreaseAcct { get; set; }
        public String PAReturnAcct { get; set; }
        public String PurchaseAcct { get; set; }
        public String PurchaseOffsetAcct { get; set; }
        public String ShippedGoodsAccount { get; set; }
        public String StockInflationOffsetAccount { get; set; }
        public String StockInflationAdjustAccount { get; set; }
        public String VATInRevenueAccount { get; set; }
        public String WHIncomingCenvatAccount { get; set; }
        public String WHOutgoingCenvatAccount { get; set; }
        public String StockInTransitAccount { get; set; }
        public String WipOffsetProfitAndLossAccount { get; set; }
        public String InventoryOffsetProfitAndLossAccount { get; set; }
        public Int32 DefaultBin { get; set; }
        public bool DefaultBinEnforced { get; set; }
        public String PurchaseBalanceAccount { get; set; }
    }

    public class ItemMasterVendors
    {
        public string BPCode { get; set; }
    }
}
