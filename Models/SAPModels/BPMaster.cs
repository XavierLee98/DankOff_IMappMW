using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class BPMaster : BusinessPartnerMaster
    {
        public List<BPAddresses> Addresses { get; set; }
        public List<BPContact> Contacts { get; set; }

        public IUserFields UserFields { get; set; }
        public BPMaster()
        {
            UserFields = new BPMasterUDF();
        }
        public class BPAddresses : BusinessPartnerMasterAddress
        {

            public IUserFields UserFields { get; set; }
            public BPAddresses()
            {
                UserFields = new BPAddressesUDF();
            }
        }

        public class BPContact : BusinesspartnerMasterContact
        {

            public IUserFields UserFields { get; set; }

            public BPContact()
            {
                UserFields = new BPContactUDF();
            }
        }
        public class BPMasterUDF : IUserFields
        {
        }
        public class BPAddressesUDF : IUserFields
        {
        }
        public class BPContactUDF : IUserFields
        {
        }
    }
}
