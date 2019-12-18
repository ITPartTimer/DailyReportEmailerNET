using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Net.Mail;
using DailyReportEmailerNET.Models;
using DailyReportEmailerNET.BLL;

namespace DailyReportEmailerNET
{
    class Program
    {
        static void Main(string[] args)
        {
            // Hold the command line argument for RptFreq
            List<string> sArgs = new List<string>();
            string brh = string.Empty;

            // Hold messages for the log file
            List<string> logMsgs = new List<string>();

            // Get command line arguments
            try
            {
                if (args.Length > 0)
                    brh = args[0].ToString();
                else
                    throw new Exception("No args in command line");
            }
            catch (Exception ex)
            {
                logMsgs.Add("Arg Exception:");
                logMsgs.Add(ex.Message.ToString());
                Logger.Log(logMsgs);

                return;
            }

            /*
            Inserting values into Excel will just append to the file.
            Need to delete the email version of Daily.xls, then
            replace with a template copy before inserting values
            */
            string fileName = "Daily.xls";
            string templatePath = @"c:\FilesToEmail\Templates";
            string destPath = @"c:\FilesToEmail\Daily";

            //File.Copy true will overwrite existing file at desination
            File.Copy(Path.Combine(templatePath, fileName), Path.Combine(destPath, fileName),true);

            /*
            Get Emails for a Brh
            */
            ReportsBLL objEmails = new ReportsBLL();

            List<string> eList = new List<string>();

            try
            {
                eList = objEmails.Get_Emails_ByBrh(brh);
            }
            catch (Exception ex)
            {
                logMsgs.Add("Get_Emails_ByBrh Exception:");
                logMsgs.Add(ex.Message.ToString());
                Logger.Log(logMsgs);

                return;
            }

            /*
            Getting MTY Bookings
            */
            ReportsBLL objBookings = new ReportsBLL();

            List<BookingsModel> bookList = new List<BookingsModel>();

            try
            {
                bookList = objBookings.Get_Bookings_MTY();
            }
            catch (Exception ex)
            {
                logMsgs.Add("Get_Employees_Reports Exception:");
                logMsgs.Add(ex.Message.ToString());
                Logger.Log(logMsgs);

                return;
            }

            /*
            Get PWC by Brh
            */
            ReportsBLL objPWC = new ReportsBLL();

            List<string> pwcList = new List<string>();

            try
            {
                pwcList = objPWC.Get_PWC_ByBrh("ST");
            }
            catch (Exception ex)
            {
                logMsgs.Add("Get_PWC_ByBrh Exception:");
                logMsgs.Add(ex.Message.ToString());
                Logger.Log(logMsgs);

                return;
            }

            /*
            Get production for each PWC and the SLT combined PWC
            NOT DOING SLT RIGHT NOW
            */
            ReportsBLL objProd = new ReportsBLL();

            List<ProdModel> pList60 = new List<ProdModel>();

            try
            {
                pList60 = objProd.Get_Prod_ByPWC("60S");
            }
            catch (Exception ex)
            {
                logMsgs.Add("Get_Prod_ByPWC(60S) Exception:");
                logMsgs.Add(ex.Message.ToString());
                Logger.Log(logMsgs);

                return;
            }

            List<ProdModel> pList72 = new List<ProdModel>();

            try
            {
                pList72 = objProd.Get_Prod_ByPWC("72S");
            }
            catch (Exception ex)
            {
                logMsgs.Add("Get_Prod_ByPWC(72S) Exception:");
                logMsgs.Add(ex.Message.ToString());
                Logger.Log(logMsgs);

                return;
            }

            List<ProdModel> pListCTL = new List<ProdModel>();

            try
            {
                pListCTL = objProd.Get_Prod_ByPWC("CTL");
            }
            catch (Exception ex)
            {
                logMsgs.Add("Get_Prod_ByPWC(CTL) Exception:");
                logMsgs.Add(ex.Message.ToString());
                Logger.Log(logMsgs);

                return;
            }

            List<ProdModel> pListMSB = new List<ProdModel>();

            try
            {
                pListMSB = objProd.Get_Prod_ByPWC("MSB");
            }
            catch (Exception ex)
            {
                logMsgs.Add("Get_Prod_ByPWC(MSB) Exception:");
                logMsgs.Add(ex.Message.ToString());
                Logger.Log(logMsgs);

                return;
            }

            /*
            Insert data into Daily.xls
            */

            // initialize text used in OleDbCommand
            string cmdText = "";

            string excelConnString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.Combine(destPath, fileName) + @";Extended Properties=""Excel 8.0;HDR=YES;""";

            using (OleDbConnection eConn = new OleDbConnection(excelConnString))
            {
                try
                {
                    eConn.Open();

                    OleDbCommand eCmd = new OleDbCommand();

                    eCmd.Connection = eConn;

                    /*
                    Tried to build the whole string outside the eConn using block, but could not get past a ; missing error
                    when inserting more than one set of values in the SQL statement.

                    Interate the list of Booking value and insert each set one-at-a-time
                    */                   

                    // Insert Bookings into Excel
                    foreach (BookingsModel b in bookList)
                    {
                        cmdText = "Insert into [Book$] (WORK_DY,BOOK_DT,SW_DLY,MS_DLY,SW_AVE,MS_AVE) Values(" + b.workDy.ToString() + "," + b.prodDt.ToString() + "," + b.swDly.ToString() + "," + b.msDly.ToString() + "," + b.swAve.ToString() + "," + b.msAve.ToString() + ");";

                        eCmd.CommandText = cmdText;

                        Console.WriteLine(cmdText);
                      
                        eCmd.ExecuteNonQuery();
                    } 

                    /*
                    Instead of hard coding each PWC, try:
                    1. Loop through PWC
                    2. cmdText tab is current PWC
                    3. Reuse <List> instead of individial for each PWC
                    4. Use parameters for the OleDb Cmd
                    */

                    /*
                    Insert production for each PWC in a specific tab
                    */                 
                    foreach (ProdModel p in pList60)
                    {
                        // PWC is a string so it needs double single quotes.
                        // Do this by adding a "'" on both sides of the property
                        cmdText = @"Insert into [60S$] (WORK_DY,PROD_DT,PWC,JOBS,LBS,BRKS,SETUPS) Values(" + p.workDy.ToString() + "," + p.prodDt.ToString() + "," + "'" + p.pwc + "'" + "," + p.jobs.ToString() + "," + p.lbs.ToString() + "," + p.brks.ToString() + "," + p.setUps.ToString() + ");";

                        eCmd.CommandText = cmdText;

                        Console.WriteLine(cmdText);

                        eCmd.ExecuteNonQuery();
                    }

                    foreach (ProdModel p in pList72)
                    {                       
                        cmdText = @"Insert into [72S$] (WORK_DY,PROD_DT,PWC,JOBS,LBS,BRKS,SETUPS) Values(" + p.workDy.ToString() + "," + p.prodDt.ToString() + "," + "'" + p.pwc + "'" + "," + p.jobs.ToString() + "," + p.lbs.ToString() + "," + p.brks.ToString() + "," + p.setUps.ToString() + ");";

                        eCmd.CommandText = cmdText;

                        Console.WriteLine(cmdText);

                        eCmd.ExecuteNonQuery();
                    }

                    foreach (ProdModel p in pListCTL)
                    {                     
                        cmdText = @"Insert into [CTL$] (WORK_DY,PROD_DT,PWC,JOBS,LBS,BRKS,SETUPS) Values(" + p.workDy.ToString() + "," + p.prodDt.ToString() + "," + "'" + p.pwc + "'" + "," + p.jobs.ToString() + "," + p.lbs.ToString() + "," + p.brks.ToString() + "," + p.setUps.ToString() + ");";

                        eCmd.CommandText = cmdText;

                        Console.WriteLine(cmdText);

                        eCmd.ExecuteNonQuery();
                    }

                    foreach (ProdModel p in pListMSB)
                    {                       
                        cmdText = @"Insert into [MSB$] (WORK_DY,PROD_DT,PWC,JOBS,LBS,BRKS,SETUPS) Values(" + p.workDy.ToString() + "," + p.prodDt.ToString() + "," + "'" + p.pwc + "'" + "," + p.jobs.ToString() + "," + p.lbs.ToString() + "," + p.brks.ToString() + "," + p.setUps.ToString() + ");";

                        eCmd.CommandText = cmdText;

                        Console.WriteLine(cmdText);

                        eCmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    logMsgs.Add("OleDb Exception:");
                    logMsgs.Add(ex.Message.ToString());
                    Logger.Log(logMsgs);

                    Console.WriteLine(ex.Message.ToString());
                }
            }

            //Email Daily.xls
            try
            {
                MailMessage mail = new MailMessage();

                //SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                SmtpClient SmtpServer = new SmtpClient("smtp.office365.com");

                //mail.From = new MailAddress("nsp.recv@gmail.com");
                mail.From = new MailAddress("sclemons@calstripsteel.com");
                mail.Subject = "Daily Reports";
                mail.Body = "Report attached";

                //Build To: line from List of emails
                foreach (string e in eList)
                {
                    logMsgs.Add(e);
                    mail.To.Add(e);
                }

                // Add attachment
                Attachment attach;
                //attach = new Attachment("c:\\FilesToEmail\\Daily.xls");
                attach = new Attachment(Path.Combine(destPath, fileName)); 
                mail.Attachments.Add(attach);

                SmtpServer.Port = 587;
                //SmtpServer.Credentials = new System.Net.NetworkCredential("nsp.recv@gmail.com", "A8dg2h8q");
                SmtpServer.Credentials = new System.Net.NetworkCredential("sclemons@calstripsteel.com", "Smet@524");
                SmtpServer.EnableSsl = true;

                //SmtpServer.Send(mail);

                Logger.Log(logMsgs);
            }
            catch (Exception ex)
            {
                logMsgs.Add("Mail Exception:");
                logMsgs.Add(ex.ToString());
                Logger.Log(logMsgs);
            }

            // testing only to stop application so I can read the console
            Console.WriteLine("Press key to exit");
            Console.ReadKey();
        }
    }
}
