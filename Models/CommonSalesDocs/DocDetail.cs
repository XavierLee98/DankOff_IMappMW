using System;

namespace IMAppSapMidware_NetCore.Models.CommonSalesDocs
{
    public class DocDetail
    {
        public int Id { get; set; }
        public int LineNum { get; set; }
        public Guid HeaderGuid { get; set; }
        public string DocNumber { get; set; }
        public string ItemCode { get; set; }
        public string ItemName { get; set; }
        public double OrderQty { get; set; }
        public double Price { get; set; }
        public double LineAmount { get; set; }
        public double TotalBeforeDiscount { get; set; }
        public double DisByPercent { get; set; }
        public double DisByValue { get; set; }
        public string TaxCode { get; set; }
        public double TaxAmt { get; set; }
        public string WhsCode { get; set; }
        public string SelectedUom { get; set; }
        public string Warehouse { get; set; }
        public DateTime DeliveryDate { get; set; }
        public string Pricelist { get; set; }
        public double TotalBeforeDis { get; set; }
        public string TaxName { get; set; }
        public double TaxRate { get; set; }
        public double GrossPrice { get; set; }
        public double LineTotal { get; set; }
        public Guid ItemGuid { get; set; }
    }
}
