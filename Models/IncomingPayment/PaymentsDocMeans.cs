using System;

namespace IMAppSapMidware_NetCore.Models.IncomingPayment
{
    public class PaymentsDocMeans
    {
        public int Id { get; set; }
        public string PaymentBy { get; set; }
        public string BankName { get; set; }
        public string ChequeNum { get; set; }
        public double Amount { get; set; }
        public Guid ItemGuid { get; set; }
        public Guid Guid { get; set; }
        public DateTime DueDate { get; set; }
        public string BankCode { get; set; }
    }
}
