using System;
using System.Collections.Generic;
using System.Text;

namespace DailyReportEmailerNET.Models
{
    public class ReportsModel
    {
        public int rptID { get; set; }
        public string name { get; set; }
        public string temppath { get; set; }
        public string rootpath { get; set; }
        public string filename { get; set; }
        public string fullpath { get; set; }
        public string frq { get; set; }
        public string typ { get; set; }
    }

    public class EmployeesReportsModel
    {
        public int rptID { get; set; }
        public string email { get; set; }
        public string name { get; set; }
        public string temppath { get; set; }
        public string rootpath { get; set; }
        public string filename { get; set; }
        public string fullpath { get; set; }
        public string frq { get; set; }
        public string typ { get; set; }
    }

    public class LogModel
    {
        public string msg { get; set; }
    }

    public class BookingsModel
    {
        public int workDy { get; set; }
        public int prodDt { get; set; }
        public int swDly { get; set; }
        public int msDly { get; set; }
        public int swAve { get; set; }
        public int msAve { get; set; }
    }

    public class ProdModel
    {
        public int workDy { get; set; }
        public int prodDt { get; set; }
        public string pwc { get; set; }
        public int jobs { get; set; }
        public int lbs { get; set; }
        public int brks { get; set; }
        public int setUps { get; set; }
    }


}
