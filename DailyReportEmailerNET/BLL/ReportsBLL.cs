using System;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DailyReportEmailerNET.DAL;
using DailyReportEmailerNET.Models;

namespace DailyReportEmailerNET.BLL
{
    [DataObject(true)]
    public class ReportsBLL
    {
        #region RPTS
        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<EmployeesReportsModel> Get_Reports_Daily_ByBrh_Name(string brh, string name)
        {
            ReportsDAL obj = new ReportsDAL();

            List<EmployeesReportsModel> lst = new List<EmployeesReportsModel>();
         
            lst = obj.LKU_Reports_Daily_ByBrh_Name(brh, name);
           
            return lst;
        }
        #endregion      

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<ProdModel> Get_Prod_ByPWC(string pwc)
        {
            ReportsDAL obj = new ReportsDAL();

            List<ProdModel> pList = new List<ProdModel>();

            pList = obj.LKU_Prod_ByPWC(pwc);

            return pList;
        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<BookingsModel> Get_Bookings_MTY_ByBrh(string brh)
        {
            ReportsDAL obj = new ReportsDAL();

            List<BookingsModel> lst = new List<BookingsModel>();

            lst = obj.LKU_Bookings_MTY_ByBrh(brh);

            return lst;
        }

        // Return list of PWC by Brh
        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<string> Get_PWC_ByBrh(string brh)
        {
            ReportsDAL obj = new ReportsDAL();

            List<string> pwcList = new List<string>();

            pwcList = obj.LKU_PWC_ByBrh(brh);

            return pwcList;
        }
    }
}
