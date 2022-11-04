using System;

namespace IMAppSapMidware_NetCore.Models.SalesEmployee
{
    public class SalesEmployee
    {
        public int SysId { get; set; }
        public string CompanyId { get; set; }
        public string UserIdName { get; set; }
        public string Password { get; set; }
        public string SapId { get; set; }
        public string DisplayName { get; set; }
        public DateTime LastModiDate { get; set; }
        public string LastModiUser { get; set; }
        public string Locked { get; set; }
        public int Roles { get; set; }
        public string PhoneNumber { get; set; }
        public string Email { get; set; }
        public Guid Assigned_token { get; set; }
        public DateTime LastLogon { get; set; }
        public string RoleDesc { get; set; }
        public int CreateERPSalesEmp { get; set; }
    }
}
