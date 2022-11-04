using Microsoft.Data.SqlClient;
using System.Data;

namespace IMAppSapMidware_NetCore.Helper.SQL
{
    public class DAC
    {
        public static string con = Program._DbMidwareConnStr; //"Server=FASTRKAPP-DEV01;Database=FT_AppMidware;User Id = sa; Password=Phan9(4$1!3#;";
        //public static string SAPcon = ConfigurationManager.ConnectionStrings["SAPConnectionString"].ToString();

        public DAC() { }
        /// <summary>
        /// Calls a stored procedure and return the result
        /// </summary>
        /// <param name="storedProcedureName">Name of the stored procedure to execute</param>
        /// <param name="arrParam">Parameters required by the stored procedure</param>
        /// <returns>DataTable containing the result</returns>
        /// <remarks></remarks>

        public static DataTable ExecuteDataTable(string storedProcedureName, params SqlParameter[] arrParam)
        {
            DataTable dt = null;
            //Open the connection
            string s = con;
            //TozDAC.Properties.Settings.Default.TozConnectionString.ToString();
            SqlConnection cnn = new SqlConnection(s);
            {
                cnn.Open();
                //Define the commands
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cnn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = storedProcedureName;
                cmd.CommandTimeout = 0;
                //Handle the parameters
                if (arrParam != null)
                {
                    foreach (SqlParameter param in arrParam)
                    {
                        if (param.Value != null)
                            cmd.Parameters.Add(param);
                    }
                }

                //Define the data adapter and fill the dataset
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                dt = new DataTable();
                da.Fill(dt);

                cnn.Close();
            }
            return dt;
        }

        public static DataTable ExecuteDataTable1(ref SqlConnection cnn, string storedProcedureName, params SqlParameter[] arrParam)
        {
            DataTable dt = null;
            //Open the connection
            string s = con;
            //TozDAC.Properties.Settings.Default.TozConnectionString.ToString();
            //SqlConnection cnn = new SqlConnection(s);
            //{
            //    cnn.Open();
            //Define the commands
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cnn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = storedProcedureName;

            //Handle the parameters
            if (arrParam != null)
            {
                foreach (SqlParameter param in arrParam)
                {
                    if (param.Value != null)
                        cmd.Parameters.Add(param);
                }
            }

            //Define the data adapter and fill the dataset
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            dt = new DataTable();
            da.Fill(dt);


            //cnn.Close();
            //}
            return dt;
        }

        public static DataTable ExecuteDataTable(string storedProceduceName)
        {
            return ExecuteDataTable(storedProceduceName, null);
        }

        /// <summary>
        /// Creates a parameter
        /// </summary>
        /// <param name="parameterName">Name of the parameter</param>
        /// <param name="parameterValue">Value of the parameter</param>
        /// <returns>SqlParameter Object</returns>
        /// <remarks>The parameter name should be the same as the property name</remarks>
        public static SqlParameter Parameter(string parameterName, object parameterValue)
        {
            SqlParameter param = new SqlParameter();
            param.ParameterName = parameterName;
            param.Value = parameterValue;
            return param;
        }

    }
}
