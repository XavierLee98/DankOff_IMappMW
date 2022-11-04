using System;

namespace IMAppSapMidware_NetCore.Models.BpUpdateGps
{
    public class BpUpdateGps
    {
        public int Id { get; set; }
        public string CardCode { get; set; }
        public string CardName { get; set; }
        public double U_Longitude { get; set; }
        public double U_Latitude { get; set; }
        public Guid Guid { get; set; }
        public int AddressType { get; set; }
        public string AddressName { get; set; }
        public string AppUser { get; set; }
        public string AppName { get; set; }
        public DateTime TransDate { get; set; }
    }
}
