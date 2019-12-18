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
        public List<ReportsModel> LKU_Reports_Active()
        {
            List<ReportsModel> lst = new List<ReportsModel>();
            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "RPT_LKU_proc_Reports_Active";
                cmd.Connection = conn;

                rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        ReportsModel r = new ReportsModel();

                        r.rptID = (int)rdr["RptID"];
                        r.name = (string)rdr["RptName"];
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

        // ---------------------------------------------------
        // Return all active employee reports with detail
        // ---------------------------------------------------
        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<EmployeesReportsModel> LKU_Employees_Reports_Active_ByBrh_Type_Freq(string brh, string typ, string frq)
        {
            List<EmployeesReportsModel> lst = new List<EmployeesReportsModel>();
           
            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);
           
            SqlConnection conn = new SqlConnection(STRATIXDataConnString);             

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "RPT_LKU_proc_Employees_Reports_Active_ByBrh_Type_Freq";
                cmd.Connection = conn;

                AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 2, ParameterDirection.Input, brh);
                AddParamToSQLCmd(cmd, "@type", SqlDbType.VarChar, 2, ParameterDirection.Input, typ);
                AddParamToSQLCmd(cmd, "@freq", SqlDbType.VarChar, 2, ParameterDirection.Input, frq);

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

        // ---------------------------------------------------
        // Return Emails for Brh
        // ---------------------------------------------------
        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<string> LKU_Emails_ByBrh( string brh)
        {
            List<string> lst = new List<string>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "RPT_LKU_proc_Employees_Active_ByBrh";
                cmd.Connection = conn;

                // Add Brh parameter
                AddParamToSQLCmd(cmd, "@brh", SqlDbType.VarChar, 2, ParameterDirection.Input, brh);

                rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    while (rdr.Read())
                    {
                        lst.Add((string)rdr["EmpEmail"]);
                    }
                }
            }
            return lst;
        }
        #endregion

        // ---------------------------------------------------
        // Get Production MTY for all PWS
        // Returns a recordset for each PWC
        // ---------------------------------------------------
        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<List<ProdModel>> LKU_Prod_All(string pwc)
        {
            List<List<ProdModel>> pList = new List<List<ProdModel>>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "ST_PROD_LKU_proc_MTY_ByPWC";
                cmd.Connection = conn;

                // Add Brh parameter
                AddParamToSQLCmd(cmd, "@pwc", SqlDbType.VarChar, 3, ParameterDirection.Input, pwc);

                rdr = cmd.ExecuteReader();

                using (rdr)
                {
                    List<ProdModel> prodList60 = new List<ProdModel>();

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

                        prodList60.Add(p);
                    }

                    pList.Add(prodList60);

                    rdr.NextResult();

                    List<ProdModel> prodList72 = new List<ProdModel>();

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

                        prodList72.Add(p);
                    }

                    pList.Add(prodList72);

                    rdr.NextResult();

                    List<ProdModel> prodListCTL = new List<ProdModel>();

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

                        prodListCTL.Add(p);
                    }

                    pList.Add(prodListCTL);

                    rdr.NextResult();

                    List<ProdModel> prodListMSB = new List<ProdModel>();

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

                        prodListMSB.Add(p);
                    }

                    pList.Add(prodListMSB);
                }
            }
            return pList;
        }

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
        public List<BookingsModel> LKU_Bookings_MTY()
        {
            List<BookingsModel> lst = new List<BookingsModel>();

            SqlCommand cmd = new SqlCommand();
            SqlDataReader rdr = default(SqlDataReader);

            SqlConnection conn = new SqlConnection(STRATIXDataConnString);

            using (conn)
            {
                conn.Open();

                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "SCORE_LKU_proc_Bookings_MTY";
                cmd.Connection = conn;

                rdr = cmd.ExecuteReader();
              
                using (rdr)
                {
                    while (rdr.Read())
                    {
                        BookingsModel r = new BookingsModel();

                        r.workDy = (int)rdr["WORK_DY"];
                        r.prodDt = (int)rdr["WORK_DY"];
                        r.swDly = (int)rdr["SW_DLY"];
                        r.msDly = (int)rdr["MS_DLY"];
                        r.swAve = (int)rdr["SW_AVE"];
                        r.msAve = (int)rdr["MS_AVE"];

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
