using System;

namespace IMAppSapMidware_NetCore.Models
{
    public class Request
    {
        public int id { get; set; }
        public string request { get; set; }
        public string sapUser { get; set; }
        public string sapPassword { get; set; }
        public DateTime requestTime { get; set; }
        public string phoneRegID { get; set; }
        public string status { get; set; }
        public Guid guid { get; set; }
        public int sapDocNumber { get; set; }
        public DateTime completedTime { get; set; }
        public int tried { get; set; }
        public int attachFileCnt { get; set; }
        public int createSAPUserSysId { get; set; }
        public string lastErrorMessage { get; set; }
    }
}
