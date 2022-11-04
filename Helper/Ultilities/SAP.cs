using IMAppSapMidware_NetCore.Models.SAPModels;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Data;

namespace IMAppSapMidware_NetCore
{
    public static class SAP
    {
        //public static SAPParam GetSAPUser(IOwinContext OwinContext)
        //{
        //    var claims = OwinContext.Authentication.User.Claims;
        //    SAPParam usr = new SAPParam();
        //    usr.UserName = claims.Where(c => c.Type == "sub").Select(c => c.Value).SingleOrDefault();
        //    usr.DBPass = claims.Where(c => c.Type == "dbpass").Select(c => c.Value).SingleOrDefault();
        //    usr.DBUser = claims.Where(c => c.Type == "dbuser").Select(c => c.Value).SingleOrDefault();
        //    usr.SAPCompany = claims.Where(c => c.Type == "sapdb").Select(c => c.Value).SingleOrDefault();
        //    usr.SAPDBType = claims.Where(c => c.Type == "servertype").Select(c => c.Value).SingleOrDefault();
        //    usr.SAPLicense = claims.Where(c => c.Type == "saplicense").Select(c => c.Value).SingleOrDefault();
        //    usr.SAPPass = claims.Where(c => c.Type == "sappass").Select(c => c.Value).SingleOrDefault();
        //    usr.SAPServer = claims.Where(c => c.Type == "sapserver").Select(c => c.Value).SingleOrDefault();
        //    usr.SAPUser = claims.Where(c => c.Type == "sapuser").Select(c => c.Value).SingleOrDefault();
        //    return usr;
        //}
        public static SAPParam GetSAPUser()
        {

            SAPParam usr = new SAPParam();


            //using (SqlConnection con = new SqlConnection("Server=FASTRKAPP-DEV01;Database=FT_AppMidware;User Id=sa;Password=Phan9(4$1!3#;"))

            using (SqlConnection con = new SqlConnection(Program._DbMidwareConnStr))
            {
                //SqlCommand cmd = new SqlCommand("select top 1 * from ft_sapsettings where sapcompany='" + sapdb + "'", con);
                SqlCommand cmd = new SqlCommand("select top 1 * from ft_sapsettings ", con);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    usr.UserName = dt.Rows[0]["UserName"].ToString();//"";
                    usr.DBPass = dt.Rows[0]["DBPass"].ToString();//"sa";
                    usr.DBUser = dt.Rows[0]["DBUser"].ToString();//"sa";
                    usr.SAPCompany = dt.Rows[0]["SAPCompany"].ToString();//"LCB_Live";
                    usr.SAPDBType = dt.Rows[0]["DBType"].ToString();//"6";// 6 is 2008, 7 is 2012
                    usr.SAPLicense = dt.Rows[0]["LicenseServer"].ToString();//"SAPSERVER:30000";
                    usr.SAPPass = dt.Rows[0]["SAPPass"].ToString();//"8899";
                    usr.SAPServer = dt.Rows[0]["Server"].ToString();//"SAPSERVER";
                    usr.SAPUser = dt.Rows[0]["SAPUser"].ToString();//"manager";
                }
            }
            return usr;
        }
        //public static SAPParam GetSAPUser()
        //{
        //    SAPParam usr = new SAPParam();
        //    return usr;
        //}
        public static List<SAPCompany> Company { get; set; }

        public static SAPCompany getSAPCompany(SAPParam sapParam)
        {
            SAPCompany com = null;

            if (Company == null) Company = new List<SAPCompany>();
            for (int i = 0; i < Company.Count; i++)
            {
                if (Company[i].UserID == sapParam.UserName)
                {
                    com = Company[i];
                    break;
                }
            }
            //if (com != null) com = null;
            //com = new SAPCompany();
            //com.UserID = sapParam.UserName;
            //com.sapParam = sapParam;
            //com.connectSAP();
            //Company.Add(com);
            if (com == null)
            {
                com = new SAPCompany();
                com.UserID = sapParam.UserName;
                com.sapParam = sapParam;
                com.connectSAP();
                Company.Add(com);
            }

            return com;
        }

    }
}
