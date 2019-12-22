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
            #region Args
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
            #endregion

            #region FileMaint
            /*
            Get Employee Reports info:
            Brh = brh
            Name = Daily
            */
            ReportsBLL objReports = new ReportsBLL();

            List<EmployeesReportsModel> empRptList = new List<EmployeesReportsModel>();

            try
            {
                empRptList = objReports.Get_Reports_Daily_ByBrh_Name(brh, "Daily");
            }
            catch (Exception ex)
            {
                logMsgs.Add("Get_Reports_Daily_ByBrh_Name Exception:");
                logMsgs.Add(ex.Message.ToString());
                Logger.Log(logMsgs);

                return;
            }

            /*
            All file related names and paths are the same for each record
            Get info from first record
            */
            string fileName = empRptList[0].filename;
            string templatePath = empRptList[0].temppath;
            string destPath = empRptList[0].rootpath;          

            //File.Copy true will overwrite existing file at desination
            File.Copy(Path.Combine(templatePath, fileName), Path.Combine(destPath, fileName),true);
            #endregion

            /*
            Getting MTY Bookings
            */
            ReportsBLL objBookings = new ReportsBLL();

            List<BookingsModel> bookList = new List<BookingsModel>();

            try
            {
                bookList = objBookings.Get_Bookings_MTY_ByBrh(brh);
            }
            catch (Exception ex)
            {
                logMsgs.Add("Get_Bookings_MTY_ByBrh Exception:");
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
                pwcList = objPWC.Get_PWC_ByBrh(brh);
            }
            catch (Exception ex)
            {
                logMsgs.Add("Get_PWC_ByBrh Exception:");
                logMsgs.Add(ex.Message.ToString());
                Logger.Log(logMsgs);

                return;
            }

            /*
            Insert data into Daily.xls
            */

            // initialize text used in OleDbCommand
            string cmdText = "";

            // Create XLS connection string
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
                        cmdText = "Insert into [Book$] (WORK_DY,BOOK_DT,BOOK_DLY,BOOK_AVE) Values(" + b.workDy.ToString() + "," + "'" + b.bookDt.ToString() + "'" + "," + b.bookDly.ToString() + "," + b.bookAve.ToString() + ");";

                        eCmd.CommandText = cmdText;

                        Console.WriteLine(cmdText);
                      
                        eCmd.ExecuteNonQuery();
                    }

                    /*
                    Build XLS Production Tabs:
                    1. Loop through PWC
                    2. cmdText tab is current PWC
                    3. Get ProdModel for PWC
                    4. Insert ProdModel into Tab for PWC
                    */
                    ReportsBLL objProd = new ReportsBLL();

                    foreach (string pwc in pwcList)
                    {
                        List<ProdModel> pwcProdList = new List<ProdModel>();

                        try
                        {
                            pwcProdList = objProd.Get_Prod_ByPWC(pwc);
                        }
                        catch (Exception ex)
                        {
                            logMsgs.Add("Get_Prod_ByPWC Exception:");
                            logMsgs.Add(ex.Message.ToString());
                            Logger.Log(logMsgs);

                            return;
                        }

                        // Loop through each record and add the XLS
                        foreach (ProdModel p in pwcProdList)
                        {
                            // PWC is a string so it needs double single quotes.
                            // Do this by adding a "'" on both sides of the property
                            cmdText = "Insert into [" + pwc + @"$] (WORK_DY,PROD_DT,PWC,JOBS,LBS,BRKS,SETUPS) Values(" + p.workDy.ToString() + "," + p.prodDt.ToString() + "," + "'" + p.pwc + "'" + "," + p.jobs.ToString() + "," + p.lbs.ToString() + "," + p.brks.ToString() + "," + p.setUps.ToString() + ");";

                            eCmd.CommandText = cmdText;

                            Console.WriteLine("Loop: " + cmdText);

                            eCmd.ExecuteNonQuery();
                        }
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

            #region Email
            /*
            Email Daily.xls to everyone at Brh who should receive it 
            */
            try
            {
                MailMessage mail = new MailMessage();

                SmtpClient SmtpServer = new SmtpClient("smtp.office365.com");

                mail.From = new MailAddress("sclemons@calstripsteel.com");
                mail.Subject = brh + " - Daily Reports";
                mail.Body = "Daily Bookings and Production attached";

                //Build To: line from emails in list of EmployeesReportsModel
                foreach (EmployeesReportsModel e in empRptList)
                {
                    logMsgs.Add(e.email.ToString());
                    mail.To.Add(e.email.ToString());
                }

                // Add attachment
                Attachment attach;
                attach = new Attachment(Path.Combine(destPath, fileName)); 
                mail.Attachments.Add(attach);

                SmtpServer.Port = 587;
                SmtpServer.Credentials = new System.Net.NetworkCredential("sclemons@calstripsteel.com", "Smet@524");
                SmtpServer.EnableSsl = true;

                SmtpServer.Send(mail);

                Logger.Log(logMsgs);
            }
            catch (Exception ex)
            {
                logMsgs.Add("Mail Exception:");
                logMsgs.Add(ex.ToString());
                Logger.Log(logMsgs);
            }
            #endregion

            // testing only to stop application so I can read the console
            Console.WriteLine("Press key to exit");
            Console.ReadKey();
        }
    }
}
