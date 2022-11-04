using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class Documents : MarketingDocument
    {
        public List<DocumentLines> Lines { get; set; }
        public List<DocumentDownPayments> DownPaymentToDraw { get; set; }

        public IUserFields UserFields { get; set; }
        public string U_PRID { get; set; }
        public string U_POID { get; set; }
        public string U_GRNID { get; set; }
        public string U_GRTNID { get; set; }
        public string U_APINVID { get; set; }
        public string U_RefNo { get; set; }
        public string U_Requester { get; set; }
        public string U_OU { get; set; }

        public Documents()
        {
            UserFields = new DocumentsUDF();
        }
    }

    public class DocumentLines : MarketingDocumentLines
    {

        public List<DocumentSerials> Serials { get; set; }
        public List<DocumentBatchs> Batches { get; set; }
        public IUserFields UserFields { get; set; }
        public string U_PCode { get; set; }
        public string U_ItemDesc { get; set; }
        public string U_BarCode { get; set; }
        public double U_Qty { get; set; }
        public string U_COLOR { get; set; }
        public double U_LENGTH { get; set; }
        public double U_WIDTH { get; set; }
        public string U_METER_CALCULATION { get; set; }
        public double U_DENSITY { get; set; }
        public double U_WEIGHT { get; set; }
        public double U_THICKNESS { get; set; }
        public string U_BOMPARENT { get; set; }

        public DocumentLines()
        {
            UserFields = new DocumentLinesUDF();
        }
    }

    public class DocumentSerials : MarketingDocumentLinesSerials
    {

        public IUserFields UserFields { get; set; }
        public DocumentSerials()
        {
            UserFields = new DocumentSerialsUDF();
        }
    }

    public class DocumentBatchs : MarketingDocumentLinesBatch
    {


        public IUserFields UserFields { get; set; }
        public DocumentBatchs()
        {
            UserFields = new DocumentBatchsUDF();
        }

    }

    public class DocumentDownPayments : MarketingDocumentDownPayment
    {


    }

    public class DocumentsUDF : IUserFields
    {
        public string Field;
    }
    public class DocumentSerialsUDF : IUserFields
    {
    }
    public class DocumentLinesUDF : IUserFields
    {
    }
    public class DocumentBatchsUDF : IUserFields
    {
    }
}
