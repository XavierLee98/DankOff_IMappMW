using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class StockTransfer : InventoryTransfer
    {
        public List<StockTransferLines> Lines { get; set; }
        public IUserFields UserFields { get; set; }
        public string U_UName { get; set; }

        public StockTransfer()
        {
            UserFields = new StockTransferUDF();
        }
    }

    public class StockTransferLines : InventoryTransferLines
    {

        public IUserFields UserFields { get; set; }

        public StockTransferLines()
        {
            UserFields = new StockTransferLinesUDF();
        }

    }
    public class StockTransferUDF : IUserFields
    {
    }
    public class StockTransferLinesUDF : IUserFields
    {
    }
}
