using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class Journal
    {
        public string LogUserID { get; set; }
        public int JournalType { get; set; }
        public DateTime ReferenceDate { get; set; }
        public String Memo { get; set; }
        public String Reference { get; set; }
        public String Reference2 { get; set; }
        public String TransactionCode { get; set; }
        public String ProjectCode { get; set; }
        public DateTime TaxDate { get; set; }
        public Int32 JdtNum { get; private set; }
        public String Indicator { get; set; }
        public bool UseAutoStorno { get; set; }
        public DateTime StornoDate { get; set; }
        public Int32 Count { get; private set; }
        public DateTime VatDate { get; set; }
        public Int32 Series { get; set; }
        public bool StampTax { get; set; }
        public DateTime DueDate { get; set; }

        private bool _AutoVAT = true;
        public bool AutoVAT { get { return _AutoVAT; } set { _AutoVAT = value; } }
        public bool ReportEU { get; set; }
        public bool Report347 { get; set; }
        public Int32 LocationCode { get; set; }
        public bool BlockDunningLetter { get; set; }
        public bool AutomaticWT { get; set; }
        public bool Corisptivi { get; set; }
        public String Reference3 { get; set; }
        public String DocumentType { get; set; }
        public bool DeferredTax { get; set; }
    }
    public class JournalLines
    {
        public String AccountCode { get; set; }
        public Double Debit { get; set; }
        public Double Credit { get; set; }
        public Double FCDebit { get; set; }
        public Double FCCredit { get; set; }
        public String FCCurrency { get; set; }
        public DateTime DueDate { get; set; }
        public String ShortName { get; set; }
        public String ContraAccount { get; set; }
        public String LineMemo { get; set; }
        public DateTime ReferenceDate1 { get; set; }
        public DateTime ReferenceDate2 { get; set; }
        public String Reference1 { get; set; }
        public String Reference2 { get; set; }
        public String ProjectCode { get; set; }
        public String CostingCode { get; set; }
        public DateTime TaxDate { get; set; }
        public Double BaseSum { get; set; }
        public String TaxGroup { get; set; }
        public Double DebitSys { get; set; }
        public Double CreditSys { get; set; }
        public DateTime VatDate { get; set; }
        public bool VatLine { get; set; }
        public Double SystemBaseAmount { get; set; }
        public Double VatAmount { get; set; }
        public Double SystemVatAmount { get; set; }
        public Double GrossValue { get; set; }
        public String AdditionalReference { get; set; }
        public Int32 CheckAbs { get; private set; }
        public String CostingCode2 { get; set; }
        public String CostingCode3 { get; set; }
        public String CostingCode4 { get; set; }
        public String TaxCode { get; set; }
        public Int32 LocationCode { get; set; }
        public String CostingCode5 { get; set; }
        public bool WTLiable { get; set; }
        public bool WTRow { get; set; }
        public String ControlAccount { get; set; }
        public bool PaymentBlock { get; set; }
        public Int32 BlockReason { get; set; }
        public String FederalTaxID { get; set; }
        public Int32 BPLID { get; set; }
    }
}
