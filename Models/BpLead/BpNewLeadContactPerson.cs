using System;

namespace IMAppSapMidware_NetCore.Models.BpLead
{
    public class BpNewLeadContactPerson
    {
        public int Id { get; set; }
        public Guid Guid { get; set; }
        public string Name { get; set; }
        public string Position { get; set; }
        public string Address { get; set; }
        public string Tel1 { get; set; }
        public string Tel2 { get; set; }
        public string Cellolar { get; set; }
        public string Fax { get; set; }
        public string E_MailL { get; set; }
        public string Pager { get; set; }
        public string BirthPlace { get; set; }
        public DateTime BirthDate { get; set; }
        public string Gender { get; set; }
        public string Profession { get; set; }
        public string Title { get; set; }
        public string BirthCity { get; set; }
        public Guid ItemGuid { get; set; }
        public string IsDefaultContact { get; set; }
    }
}
