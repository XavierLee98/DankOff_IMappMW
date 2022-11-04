using System;

namespace IMAppSapMidware_NetCore.Models.IncomingPayment
{
    public class PaymentsDocDetails
    {
        //public int Id { get; set; }
        //public int DocNum { get; set; }
        //public DateTime DocDate { get; set; }
        //public DateTime DueDate { get; set; }
        //public decimal DocAmt { get; set; }
        //public decimal DiscountByPercent { get; set; }
        //public double PaymentAmount { get; set; }
        //public Guid ItemGuid { get; set; }
        //public Guid Guid { get; set; }
        //public string DocType { get; set; }
        //public int DocEntry { get; set; }
        //public int InstllmntId { get; set; }
        public int Id { get; set; }
        public int DocEntry { get; set; }
        public string DocNum { get; set; }
        public DateTime DocDate { get; set; }
        public DateTime DueDate { get; set; }
        public double DocAmt { get; set; } // local currency
        public double DocAmtFC { get; set; } // foreign currency 20200312T1627
        public double DiscountByPercent { get; set; }
        public double PaymentAmount { get; set; } // local currency, use in app
        public double PaymentAmountFC { get; set; } // foreign currency 20200312T1627, use in app
        public Guid ItemGuid { get; set; }
        public Guid Guid { get; set; }
        public string DocType { get; set; }

        // 20200305T1921
        
        public int InvDocEntry { get; set; } // keep the invoice doc entry
        public int InstlmntID { get; set; } // for inv installment entry

        // 20200312T1626
        public string Currency { get; set; } // represent the local currency, will get from the app local database
        public string CurrencyFC { get; set; } // represent the foreignt currency, 
    }
}
