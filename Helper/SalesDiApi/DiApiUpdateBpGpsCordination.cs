using IMAppSapMidware_NetCore.Models.BpUpdateGps;
using SAPbobsCOM;
using System;

namespace IMAppSapMidware_NetCore.Helper.DiApi
{
    public class DiApiUpdateBpGpsCordination
    {
        public BpUpdateGps BpUpdate { get; set; } = null;
        public ErpPropertyHelper ErpProperty { get; set; } = null;
        public string LastSAPMsg { get; set; } = string.Empty;
        public string Midware_DBConnStr { get; set; } = string.Empty;
        public string Erp_DBConnStr { get; set; } = string.Empty;
        public string Guid { get; set; } = string.Empty;
        public string DocType { get; set; } = string.Empty;

        private Company sapCompany { get; set; }

        public bool UpdateBpAddressGps()
        {
            string modName = "[UpdateBpAddressGps][" + DocType + "]";
            try
            {
                if (BpUpdate == null)
                {
                    Log($"{modName}\nBp GPS update object empty.");
                    return false;
                }

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
                
                var bp = (BusinessPartners)sapCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                var isBp = bp.GetByKey(BpUpdate.CardCode);

                if (!isBp)
                {
                    Log($"{modName}\n{sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
                    return false;
                }

                if (bp.Addresses == null)
                {
                    Log($"{modName}\n{sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
                    return false;
                }

                var addressType = (BpUpdate.AddressType == 0) ? BoAddressType.bo_ShipTo : BoAddressType.bo_BillTo;

                for (int x = 0; x < bp.Addresses.Count; x++)
                {
                    bp.Addresses.SetCurrentLine(x);
                    if (bp.Addresses.AddressName.Equals(BpUpdate.AddressName) &&
                        bp.Addresses.AddressType.Equals(addressType))
                    {
                        bp.Addresses.UserFields.Fields.Item(nameof(BpUpdate.U_Longitude)).Value = BpUpdate.U_Longitude;
                        bp.Addresses.UserFields.Fields.Item(nameof(BpUpdate.U_Latitude)).Value = BpUpdate.U_Latitude;

                        int result = bp.Update();
                        if (result == 0)
                        {
                            // process the success hanlder
                            Log($"{modName} GPS Updated, Bp: {BpUpdate.CardCode}, {BpUpdate.CardName} , {Guid}");
                            return true; // exit the code
                        }
                    }
                    // continue the for loop
                }

                // process the failure handler
                // else                
                Log($"{modName}\n{sapCompany.GetLastErrorCode()}\n{sapCompany.GetLastErrorDescription()}");
                return false;
            }
            catch (Exception e)
            {
                Log($"{e.Message}\n{e.StackTrace}");
                return false;
            }
        }

        void Log(string message)
        {
            LastSAPMsg += $"\n\n{message}";
            Program.FilLogger?.Log(message);
        }
    }
}
