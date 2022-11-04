using System;

namespace IMAppSapMidware_NetCore.Models.BpLead
{
    public class BpNewLeadShipBillAddr
    {
        public int Id { get; set; }
        public Guid Guid { get; set; }
        public string Address { get; set; }
        public string Street { get; set; }
        public string Block { get; set; }
        public string ZipCode { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
        public string State { get; set; }
        public string Building { get; set; }
        public string StreetNo { get; set; }
        public string AdresType { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public Guid ItemGuid { get; set; }
        public string IsDefaultAddrs { get; set; }
    }
}
