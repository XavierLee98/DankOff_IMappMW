using IMAppSapMidware_NetCore.Helper.DiApi;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Helper.SalesDiApi
{
    public class DiApiUpdatePickList
    {
        public ErpPropertyHelper ErpProperty { get; set; } = null;
        public string PostedDocNum { get; set; } = string.Empty;
        public string LastSAPMsg { get; set; } = string.Empty;
        public string Midware_DBConnStr { get; set; } = string.Empty;
        public string Erp_DBConnStr { get; set; } = string.Empty;
        public string Guid { get; set; } = string.Empty;

        Company sapCompany { get; set; }

        void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }

        //Update Pick List and Partially Pick Items in Pick List
        public bool PartialUpdatePickList()
        {
            string modName = $"[UpdatePickList]";

            try
            {
                bool retResult = false;
                // connect the diapi company from class
                // maintain in one place
                // replace if (ConnectDI() != 0) return false;  // connect the server if fail then return                                
                var di = new DiApiUtilities { ErpProperty = this.ErpProperty };
                if (di.ConnectDI() != 0)
                {
                    Log($"{modName}\n{di.LastErrorMessage}");
                    return false;
                }
                sapCompany = di.SapCompany;

                SAPbobsCOM.PickLists oPickLists = null;
                SAPbobsCOM.PickLists_Lines oPickLists_Lines = null;
                oPickLists = (SAPbobsCOM.PickLists)sapCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPickLists);
                oPickLists.GetByKey(2);
                oPickLists_Lines = oPickLists.Lines;

                // When Item on Row 1 has 10 Quantity and You want to Pick Only 3 Quantity
                oPickLists_Lines.SetCurrentLine(0);
                oPickLists_Lines.PickedQuantity = 3;
                //Pick 1 Quantity of Batch "B1_001" for Item on Row 1
                oPickLists_Lines.BatchNumbers.BatchNumber = "B1_001";
                oPickLists_Lines.BatchNumbers.Quantity = 1;
                oPickLists_Lines.BatchNumbers.BaseLineNumber = 0; //BaseLineNumber is Row Number on PKL1

                //Pick 2 Quantity of Batch "B1_002" for Item on Row 1
                oPickLists_Lines.BatchNumbers.Add();
                oPickLists_Lines.BatchNumbers.BatchNumber = "B1_002";
                oPickLists_Lines.BatchNumbers.Quantity = 2;
                oPickLists_Lines.BatchNumbers.BaseLineNumber = 0;

                // When Item on Row 2 has 20 Quantity and You want to Pick Only 4 Quantity
                oPickLists_Lines.SetCurrentLine(1);
                oPickLists_Lines.PickedQuantity = 4;
                //Pick 3 Quantity of Batch "B2_001" for Item on Row 2
                oPickLists_Lines.BatchNumbers.BatchNumber = "B2_001";
                oPickLists_Lines.BatchNumbers.Quantity = 3;
                oPickLists_Lines.BatchNumbers.BaseLineNumber = 1;

                //Pick 1 Quantity of Batch "B2_002" for Item on Row 2
                oPickLists_Lines.BatchNumbers.Add();
                oPickLists_Lines.BatchNumbers.BatchNumber = "B2_002";
                oPickLists_Lines.BatchNumbers.Quantity = 1;
                oPickLists_Lines.BatchNumbers.BaseLineNumber = 1; //BaseLineNumber is Row Number on PKL1
                int RetVal = oPickLists.Update();
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
            }

            return true;
        }


    }
}
