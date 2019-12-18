using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DailyReportEmailerNET.DAL
{
    public class SQLHelpers
    {
        private string _STRATIXDataConnString;    

        // --------------------------------------------
        // Database connection string
        // --------------------------------------------
        public string STRATIXDataConnString
        {
            get
            {
                _STRATIXDataConnString = "Data Source=THINKM81\\SQL2014DEV;Initial Catalog=STRATIXData;User ID=ReportRunner;Password=vH1bC1J6W";
                //_STRATIXDataConnString = "Data Source=VC-030\\SQLDEV2014VC030;Initial Catalog=STRATIXData;User ID=ReportRunner;Password=vH1bC1J6W";

                if (string.IsNullOrEmpty(_STRATIXDataConnString))
                {
                    throw new NullReferenceException("Empty Connection String");
                }
                else
                {
                    return _STRATIXDataConnString;
                }
            
            }

        } //STRATIXDataConnString

        // ---------------------------------------------------------
        // Add input parameters to a SQL command object.  Overloaded
        // to accept a param value for input parameters
        // ---------------------------------------------------------
        protected void AddParamToSQLCmd(SqlCommand sqlCmd, string paramID,
                                        SqlDbType sqlType, int paramSize,
                                        ParameterDirection paramDirection, object paramValue)
        {
            //See VB code for some error checking to insert here

            SqlParameter newSqlParam = new SqlParameter();

            newSqlParam.ParameterName = paramID;
            newSqlParam.SqlDbType = sqlType;
            newSqlParam.Size = paramSize;
            newSqlParam.Direction = paramDirection;
            newSqlParam.Value = paramValue;

            sqlCmd.Parameters.Add(newSqlParam);
        }

        // ------------------------------------------------------
        // Overload for input parameter to allow proper config.
        // of a Decimal SqlDbType paramSize=Precision, scale=Scale
        // -------------------------------------------------------
        protected void AddParamToSQLCmd(SqlCommand sqlCmd, string paramID,
                                        SqlDbType sqlType, int paramSize, int scale,
                                        ParameterDirection paramDirection, object paramValue)
        {
            //See VB code for some error checking to insert here

            SqlParameter newSqlParam = new SqlParameter();

            newSqlParam.ParameterName = paramID;
            newSqlParam.SqlDbType = sqlType;
            newSqlParam.Precision = (byte)paramSize;
            newSqlParam.Scale = (byte)scale;
            newSqlParam.Direction = paramDirection;
            newSqlParam.Value = paramValue;

            sqlCmd.Parameters.Add(newSqlParam);
        }

        // ---------------------------------------------------
        // Add output parameters to SQL comment object
        // ---------------------------------------------------
        protected void AddParamToSQLCmd(SqlCommand sqlCmd, string paramID,
                                        SqlDbType sqlType, int paramSize,
                                        ParameterDirection paramDirection)
        {
            //See VB code for some error checking to insert here

            SqlParameter newSqlParam = new SqlParameter();

            newSqlParam.ParameterName = paramID;
            newSqlParam.SqlDbType = sqlType;
            newSqlParam.Size = paramSize;
            newSqlParam.Direction = paramDirection;

            sqlCmd.Parameters.Add(newSqlParam);
        }
    }
}
