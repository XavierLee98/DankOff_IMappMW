using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class BankingPayment
    {
        public string LogUserID { get; set; }
        public Int32 DocNum { get; set; }
        public string DocObject { get; set; }
        public int DocType { get; set; }
        public bool HandWritten { get; set; }
        public bool Printed { get; private set; }
        public DateTime DocDate { get; set; }
        public String CardCode { get; set; }
        public String CardName { get; set; }
        public String Address { get; set; }
        public String CashAccount { get; set; }
        public String DocCurrency { get; set; }
        public Double CashSum { get; set; }
        public String CheckAccount { get; set; }
        public String TransferAccount { get; set; }
        public Double TransferSum { get; set; }
        public DateTime TransferDate { get; set; }
        public String TransferReference { get; set; }
        public bool LocalCurrency { get; set; }
        public Double DocRate { get; set; }
        public String Reference1 { get; set; }
        public String Reference2 { get; set; }
        public String CounterReference { get; set; }
        public String Remarks { get; set; }
        public String JournalRemarks { get; set; }
        public Int32 ContactPersonCode { get; set; }
        public bool ApplyVAT { get; set; }
        public DateTime TaxDate { get; set; }
        public Int32 Series { get; set; }
        public String BankCode { get; set; }
        public String BankAccount { get; set; }
        public String ProjectCode { get; set; }
        public Double DeductionPercent { get; set; }
        public Double DeductionSum { get; set; }
        public Double CashSumFC { get; private set; }
        public Double CashSumSys { get; private set; }
        public String BoeAccount { get; set; }
        public Double BillOfExchangeAmount { get; set; }
        public Double BillOfExchangeAmountFC { get; private set; }
        public Double BillOfExchangeAmountSC { get; private set; }
        public String BillOfExchangeAgent { get; set; }
        public String WTCode { get; set; }
        public Double WTAmount { get; set; }
        public Double WTAmountFC { get; private set; }
        public Double WTAmountSC { get; private set; }
        public String WTAccount { get; private set; }
        public Double WTTaxableAmount { get; private set; }
        public bool Proforma { get; set; }
        public String TaxGroup { get; set; }
        public String PayToBankAccountNo { get; set; }
        public String PayToCode { get; set; }
        public String PayToBankCountry { get; set; }
        public bool IsPayToBank { get; set; }
        public Int32 DocEntry { get; private set; }
        public String PayToBankCode { get; set; }
        public String PayToBankBranch { get; set; }
        public Double BankChargeAmount { get; set; }
        public Double BankChargeAmountInFC { get; private set; }
        public Double BankChargeAmountInSC { get; private set; }
        public DateTime VatDate { get; set; }
        public String TransactionCode { get; set; }
        public Double TransferRealAmount { get; set; }
        public Double UnderOverpaymentdifference { get; private set; }
        public Double UnderOverpaymentdiffSC { get; private set; }
        public Double WtBaseSum { get; set; }
        public Double WtBaseSumFC { get; private set; }
        public Double WtBaseSumSC { get; private set; }
        public DateTime DueDate { get; set; }
        public Int32 LocationCode { get; set; }
        public bool Cancelled { get; private set; }
        public String ControlAccount { get; set; }
        public Double UnderOverpaymentdiffFC { get; private set; }
        public Int32 BPLID { get; set; }
        public String BPLName { get; private set; }
        public String VATRegNum { get; private set; }
    }

    public class BankingPaymentCheque
    {
        public Int32 LineNum { get; private set; }
        public DateTime DueDate { get; set; }
        public Int32 CheckNumber { get; set; }
        public String BankCode { get; set; }
        public String Branch { get; set; }
        public String AccounttNum { get; set; }
        public String Details { get; set; }
        public bool Trnsfrable { get; set; }
        public Double CheckSum { get; set; }
        public String CountryCode { get; set; }
        public Int32 CheckAbsEntry { get; private set; }
        public String CheckAccount { get; set; }
        public bool ManualCheck { get; set; }
    }

    public class BankingPaymentAccounts
    {
        public Int32 LineNum { get; private set; }
        public String AccountCode { get; set; }
        public Double SumPaid { get; set; }
        public String Decription { get; set; }
        public String VatGroup { get; set; }
        public String AccountName { get; set; }
        public Double GrossAmount { get; set; }
        public String ProfitCenter { get; set; }
        public String ProjectCode { get; set; }
        public Double VatAmount { get; set; }
        public Int32 LocationCode { get; private set; }
        public String ProfitCenter2 { get; set; }
        public String ProfitCenter3 { get; set; }
        public String ProfitCenter4 { get; set; }
        public String ProfitCenter5 { get; set; }
        public Double EqualizationVatAmount { get; private set; }
    }
    public class BankingPaymentInvoices
    {
        public Int32 LineNum { get; private set; }
        public Int32 DocEntry { get; set; }
        public int InvoiceType { get; set; }
        public Double SumApplied { get; set; }
        public Double AppliedFC { get; set; }
        public Int32 DocLine { get; set; }
        public Double DiscountPercent { get; set; }
        public Double PaidSum { get; private set; }
        public Int32 InstallmentId { get; set; }
        public Double WitholdingTaxApplied { get; private set; }
        public Double WitholdingTaxAppliedFC { get; private set; }
        public Double WitholdingTaxAppliedSC { get; private set; }
        public DateTime LinkDate { get; private set; }
        public String DistributionRule { get; set; }
        public String DistributionRule2 { get; set; }
        public String DistributionRule3 { get; set; }
        public String DistributionRule4 { get; set; }
        public String DistributionRule5 { get; set; }
        public Double TotalDiscount { get; set; }
        public Double TotalDiscountFC { get; set; }
        public Double TotalDiscountSC { get; private set; }
    }

    public class BankingPaymentFormItems
    {
        public Int32 CashFlowAssignmentsID { get; private set; }
        public Int32 CashFlowLineItemID { get; set; }
        public String CheckNumber { get; set; }
        public Double AmountLC { get; set; }
        public Double AmountFC { get; set; }

        public int PaymentMeans { get; set; }
    }
}
