using APAR_ManpowerRequisition_Mail_Schedular.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UserInformation;
namespace APAR_ManpowerRequisition_Schedular.Models
{
    public static class SQLUtility
    {
        static UserOperation _UserOperation = new UserOperation();
        private static SqlConnection GetSQLConnection()
        {
            //string _ServerName = CustomSharePointUtility.Decrypt(_UserOperation.ReadValue("Server_Name_Live"));
            string _ServerName = _UserOperation.ReadValue("Server_Name_Live");
            string _DBName = CustomSharePointUtility.Decrypt(_UserOperation.ReadValue("DB_Name_Live"));
            string _DBUid = CustomSharePointUtility.Decrypt(_UserOperation.ReadValue("DB_USER_ID_Live"));
            string _DBPassword = CustomSharePointUtility.Decrypt(_UserOperation.ReadValue("DB_PAssword_Live"));

            string connectionString = "Data Source=" + _ServerName + ";Initial Catalog=" + _DBName + ";user id=" + _DBUid + ";password=" + _DBPassword;
          
            return new SqlConnection(connectionString);
        }
        public static DataTable GetDataTable( string sqlQuery)
        {
            DataTable _returnTable = null;

            using (SqlConnection _connection = GetSQLConnection())
            {


                SqlCommand _sqlcmd = null;
                try
                {

                    _connection.Open();
                    _sqlcmd = _connection.CreateCommand();
                    _sqlcmd.CommandType = CommandType.Text;
                    _sqlcmd.CommandText = sqlQuery;
                    _sqlcmd.ExecuteNonQuery();
                    SqlDataAdapter odata = new SqlDataAdapter(_sqlcmd);
                    _returnTable = new DataTable();
                    odata.Fill(_returnTable);

                    CustomSharePointUtility.WriteLog("Connected to SQL");

                }
                catch (Exception ex)
                {
                    _returnTable = null;
                    throw ex;
                }
                finally
                {
                    if (_sqlcmd != null)
                    {
                        _sqlcmd = null;
                    }

                }

            }

            return _returnTable;
        }

        #region ReadingSQLquery
        public static string ReadQuery(string QueryFileName)
        {
            try
            {
                System.IO.StreamReader myFile =
                new System.IO.StreamReader(QueryFileName);
                string query = myFile.ReadToEnd();

                myFile.Close();
                return query;

            }
            catch (Exception ex)
            {
                CustomSharePointUtility.WriteLog(ex.ToString());
                return "";

            }
        }
        #endregion
    }
}
