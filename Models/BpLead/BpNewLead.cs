using System;

namespace IMAppSapMidware_NetCore.Models.BpLead
{
    public class BpNewLead
    {
        public int Id { get; set; }
        public string Cardcode { get; set; }
        public string CardName { get; set; }
        public string CardFName { get; set; }
        public string Currency { get; set; }
        public string Phone1 { get; set; }
        public string Phone2 { get; set; }
        public string Cellular { get; set; }
        public string Fax { get; set; }
        public string E_Mail { get; set; }
        public string IntrntSite { get; set; }
        public string RegNum { get; set; }
        public string SlpCode { get; set; }
        public Guid Guid { get; set; }
        public DateTime RegDate { get; set; }
        public DateTime TransDate { get; set; }
        public string ShippingType { get; set; }
    }
}
