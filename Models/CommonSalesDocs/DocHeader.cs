using System;

namespace IMAppSapMidware_NetCore.Models.CommonSalesDocs
{
    public class DocHeader
    {
        public int Id { get; set; }
        public string DocType { get; set; }
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public double DocAmt { get; set; }
        public string DocNum { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime TaxDate { get; set; }
        public DateTime DueDate { get; set; }
        public string Ref2 { get; set; }
        public string Comments { get; set; }
        public string JrnlMemo { get; set; }
        public string NumberAtCard { get; set; }
        public Guid Guid { get; set; }
        public bool Submitted { get; set; }
        public string SapDocNo { get; set; }
        public string ContactPerson { get; set; }
        public string ShipAddress { get; set; }
        public string BillAddress { get; set; }
        public double DiscountByPercent { get; set; }
        public double DiscountByValue { get; set; }
        public double TaxAmt { get; set; }
        public int NumberFileAttached { get; set; }
        public string DocSeries { get; set; }
        public string Currency { get; set; }
        public string SalesPerson { get; set; }
        public double TotalBeforeDis { get; set; }
        public bool IsPostDraft { get; set; }
    }
}
