using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class SAPParam
    {
        public string UserName { get; set; }
        public string SAPDBType { get; set; }
        public string SAPServer { get; set; }
        public string DBUser { get; set; }
        public string DBPass { get; set; }
        public string SAPCompany { get; set; }
        public string SAPUser { get; set; }
        public string SAPPass { get; set; }
        public string SAPLicense { get; set; }
    }
}
