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
        // NOT IN USE - 12/15/19
        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<ReportsModel> Get_Reports()
        {
            ReportsDAL obj = new ReportsDAL();

            List<ReportsModel> lst = new List<ReportsModel>();
          
            lst = obj.LKU_Reports_Active();         

            return lst;
        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<EmployeesReportsModel> Get_Employees_Reports_ByBrh_Type_Freq(string brh, string typ, string frq)
        {
            ReportsDAL obj = new ReportsDAL();

            List<EmployeesReportsModel> lst = new List<EmployeesReportsModel>();
         
            lst = obj.LKU_Employees_Reports_Active_ByBrh_Type_Freq(brh, typ, frq);
           
            return lst;
        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<string> Get_Emails_ByBrh(string brh)
        {
            ReportsDAL obj = new ReportsDAL();

            List<string> lst = new List<string>();

            lst = obj.LKU_Emails_ByBrh(brh);

            return lst;
        }
        #endregion

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<List<ProdModel>> Get_Prod_All(string pwc)
        {
            ReportsDAL obj = new ReportsDAL();

            List<List<ProdModel>> pList = new List<List<ProdModel>>();

            pList =  obj.LKU_Prod_All(pwc);

            return pList;
        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<ProdModel> Get_Prod_ByPWC(string pwc)
        {
            ReportsDAL obj = new ReportsDAL();

            List<ProdModel> pList = new List<ProdModel>();

            pList = obj.LKU_Prod_ByPWC(pwc);

            return pList;
        }

        [DataObjectMethod(DataObjectMethodType.Select)]
        public List<BookingsModel> Get_Bookings_MTY()
        {
            ReportsDAL obj = new ReportsDAL();

            List<BookingsModel> lst = new List<BookingsModel>();

            lst = obj.LKU_Bookings_MTY();

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
