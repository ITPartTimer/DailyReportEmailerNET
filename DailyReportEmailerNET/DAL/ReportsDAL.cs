using System;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DailyReportEmailerNET.Models;

namespace DailyReportEmailerNET.DAL
{
    [DataObject(true)]
    public class ReportsDAL : SQLHelpers
    {
        #region RPTS
        // ---------------------------------------------------
        // Return all active reports
        // ---------------------------------------------------
        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<EmployeesReportsModel> LKU_Reports_Daily_ByBrh_Name(string brh, string name)
        {
            List<EmployeesReportsModel> lst = new List<EmployeesReportsModel>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "RPT_LKU_proc_Daily_ByBrh_Name";
                cmd.Connection = conn;

                AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 2, ParameterDirection.Input, brh);
                AddParamToSQLCmd(cmd, "@name", SqlDbType.VarChar, 25, ParameterDirection.Input, name);

                rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        EmployeesReportsModel r = new EmployeesReportsModel();

                        r.rptID = (int)rdr["RptID"];
                        r.email = (string)rdr["EmpEmail"];
                        r.name = (string)rdr["RptName"];
                        r.temppath = (string)rdr["RptTempPath"];
                        r.rootpath = (string)rdr["RptRootPath"];
                        r.filename = (string)rdr["RptFileName"];
                        r.fullpath = (string)rdr["RptFullPath"];
                        r.frq = (string)rdr["RptFreq"];
                        r.typ = (string)rdr["RptType"];

                        lst.Add(r);
                    }
                }
            }
            
            return lst;
        }     
        #endregion

        // ---------------------------------------------------
        // Get Production MTY by PWC
        // ---------------------------------------------------
        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<ProdModel> LKU_Prod_ByPWC(string pwc)
        {
            List<ProdModel> pList = new List<ProdModel>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "ST_PROD_LKU_proc_MTY_ByPWC";
                cmd.Connection = conn;

                AddParamToSQLCmd(cmd, "@pwc", SqlDbType.VarChar, 3, ParameterDirection.Input, pwc);

                rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        ProdModel p = new ProdModel();

                        p.workDy = (int)rdr["WORK_DY"];
                        p.prodDt = (int)rdr["PROD_DT"];
                        p.pwc = (string)rdr["PWC"];
                        p.jobs = (int)rdr["JOBS"];
                        p.lbs = (int)rdr["LBS"];
                        p.brks = (int)rdr["BRKS"];
                        p.setUps = (int)rdr["SETUPS"];

                        pList.Add(p);
                    }
                }
            }
            return pList;
        }

        // ---------------------------------------------------
        // Return Bookings MTY
        // ---------------------------------------------------
        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<BookingsModel> LKU_Bookings_MTY_ByBrh(string brh)
        {
            List<BookingsModel> lst = new List<BookingsModel>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "RPT_LKU_proc_Bookings_MTY_ByBrh";
                cmd.Connection = conn;

                AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 2, ParameterDirection.Input, brh);

                rdr = cmd.ExecuteReader();
              
                using (rdr)
                {
                    while (rdr.Read())
                    {
                        BookingsModel r = new BookingsModel();
                      
                        r.workDy = (int)rdr["WORK_DY"];
                        r.bookDt = (DateTime)rdr["BOOK_DT"];
                        r.bookDly = (int)rdr["BOOK_DLY"];
                        r.bookAve = (int)rdr["BOOK_AVE"];

                        lst.Add(r);
                    }
                }
            }
            return lst;
        }

        // ---------------------------------------------------
        // Return Active PWC by Brh
        // ---------------------------------------------------
        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<string> LKU_PWC_ByBrh(string brh)
        {
            List<string> pwcList = new List<string>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "PWC_LKU_proc_Active_ByBrh";
                cmd.Connection = conn;

                AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 3, ParameterDirection.Input, brh);

                rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        string s;

                        s = (string)rdr["PWC"];

                        pwcList.Add(s);  
                    }
                }
            }
            return pwcList;
        }
    }
}
