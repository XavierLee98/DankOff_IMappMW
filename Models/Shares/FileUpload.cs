using System;

namespace IMAppSapMidware_NetCore.Models.Share
{
    public class FileUpload
    {
        public int Id { get; set; }
        public Guid HeaderGuid { get; set; }
        public DateTime UploadDatetime { get; set; }
        public string AppUser { get; set; }
        public string ServerSavedPath { get; set; }
    }
}
