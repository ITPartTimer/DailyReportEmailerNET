using System;
using System.ComponentModel;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DailyReportEmailerNET.BLL
{
    [DataObjectAttribute]
    public class Logger
    {
        [DataObjectMethodAttribute(DataObjectMethodType.Insert)]
        public static void Log(List<string> msgs)
        {
            string ReportPath = @"C:\FilesToEmail\Log.txt";

            using (StreamWriter writer = new StreamWriter(ReportPath, true))
            {
                writer.WriteLine("-------- DailyReportMailerNET --------");
                writer.WriteLine("Date: " + DateTime.Now.ToString());

                foreach(string m in msgs)
                {
                    writer.WriteLine(m);
                }              
            }
        }
    }
}
