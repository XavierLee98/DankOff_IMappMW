using System;
using System.Collections.Generic;
using System.Text;

namespace IMAppSapMidware_NetCore.Models.SAPModels
{
    public class Items : ItemMaster
    {
        //public List<ItemsPrices> PriceList { get; set; }
        public List<ItemsWhsInfo> WhsInfo { get; set; }
        public List<ItemsVendor> PreferedVendor { get; set; }
        public IUserFields UserFields { get; set; }

        public Items()
        {
            UserFields = new ItemsUDF();
        }
    }

    public class ItemsPrices : ItemMasterPriceList
    {

        public IUserFields UserFields { get; set; }

        public ItemsPrices()
        {
            UserFields = new ItemsPricesUDF();
        }
    }
    public class ItemsWhsInfo : ItemMasterWarehouse
    {


        public IUserFields UserFields { get; set; }

        public ItemsWhsInfo()
        {
            UserFields = new ItemsWhsInfoUDF();
        }
    }

    public class ItemsVendor : ItemMasterVendors
    {

    }

    public class ItemsUDF : IUserFields
    {
        public string U_Test { get; set; }
    }
    public class ItemsPricesUDF : IUserFields
    {
    }
    public class ItemsWhsInfoUDF : IUserFields
    {
    }
}
