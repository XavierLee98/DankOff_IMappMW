using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class Payments : BankingPayment
    {

        public List<PaymentsInvoices> AppliedInvoices { get; set; }
        public List<PaymentsCheques> ChequesPayment { get; set; }
        public List<PaymentsAccounts> AccountsPayment { get; set; }
        public List<PaymentsFormItems> FormItems { get; set; }
        public IUserFields UserFields { get; set; }

        public string U_PAYID { get; set; }


        public Payments()
        {
            UserFields = new PaymentsUDF();
        }
    }

    public class PaymentsCheques : BankingPaymentCheque
    {

        public IUserFields UserFields { get; set; }

        public PaymentsCheques()
        {
            UserFields = new PaymentsChequesUDF();
        }
    }
    public class PaymentsAccounts : BankingPaymentAccounts
    {
        public IUserFields UserFields { get; set; }

        public PaymentsAccounts()
        {
            UserFields = new PaymentsAccountsUDF();
        }

    }
    public class PaymentsInvoices : BankingPaymentInvoices
    {

        public IUserFields UserFields { get; set; }

        public PaymentsInvoices()
        {
            UserFields = new PaymentsInvoicesUDF();
        }
    }

    public class PaymentsFormItems : BankingPaymentFormItems
    {



    }
    public class PaymentsUDF : IUserFields
    {
    }
    public class PaymentsInvoicesUDF : IUserFields
    {
    }
    public class PaymentsFormItemsUDF : IUserFields
    {
    }
    public class PaymentsChequesUDF : IUserFields
    {
    }
    public class PaymentsAccountsUDF : IUserFields
    {
    }
}
