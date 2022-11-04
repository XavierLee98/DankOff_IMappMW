using System;

namespace IMAppSapMidware_NetCore.Models.IncomingPayment
{
    public class PaymentsDocHeader
    {
        public string DocType { get; set; }
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public decimal DocAmt { get; set; }
        public string DocNum { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime TaxDate { get; set; }
        public DateTime DueDate { get; set; }
        public string Ref2 { get; set; }
        public string Comments { get; set; }
        public string JrnlMemo { get; set; }
        public int Id { get; set; }
        public Guid Guid { get; set; }
        public short Submitted { get; set; }
        public string SapDocNo { get; set; }
        public short IsPostDraft { get; set; }
        public string DocStatus { get; set; }
        public DateTime TransDate { get; set; }
        public double CollectedCash { get; set; }
        public string TransferAccount { get; set; }
        public DateTime TransferDate { get; set; }
        public double TransferSum { get; set; }
        public string TransferReference { get; set; }
    }
}
