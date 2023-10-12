using Quartz;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using Quartz.Core;
using System.Threading.Tasks;
using System.Data;
using System.Globalization;
using System.Runtime.InteropServices;
using System.IO;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Collections;
using System.Diagnostics;

namespace CronService_Processor
{
    public class ScheduledJobs : IJob
    {
        EventLog EventLog = new EventLog();

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        public void Execute(IJobExecutionContext context)
        {
            try
            {
                EventLog.Log = "CronService_Processor";
                if (!EventLog.SourceExists("Cron Email Service"))
                {
                    EventLog.CreateEventSource("Cron Email Service", "Cron Email Service");
                }
                EventLog.WriteEntry("Cron Email Service","Starting Email Service",EventLogEntryType.Information);
                string[] reports = ConfigurationManager.AppSettings["Reports"].ToString().Split(',');
                try
                {
                    Access access = new Access();
                    foreach (string reportName in reports)
                    {
                        EventLog.WriteEntry("Cron Email Service", $"fetching Report Details for {reportName}", EventLogEntryType.Information);
                        string filePath = System.IO.Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory + "/Reports/" + reportName + "/" + string.Format("{0}.txt", reportName));
                        string cmd = File.ReadAllText(filePath);
                        DataTable dt = access.GetTable(cmd);
                        if (dt.Rows.Count > 0 && !string.IsNullOrEmpty(reportName))
                        {
                            string filepath = System.IO.Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory + "/Reports/" + reportName + "/" + string.Format("{0}.xlsx", reportName));
                            if (File.Exists(filepath))
                            {
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                File.Delete(filepath);
                            }
                            EventLog.WriteEntry("Cron Email Service", $"Exporting {reportName} to Excel", EventLogEntryType.Information);
                            string filename = ExportToExcel(dt, reportName);
                            EventLog.WriteEntry("Cron Email Service", $"Sending Email for {reportName}", EventLogEntryType.Information);
                            SendEmail(filename, reportName);
                            EventLog.WriteEntry("Cron Email Service", $"{reportName} Email Sent", EventLogEntryType.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    EventLog.WriteEntry("Cron Email Service", $"Exception occurred - {ex.Message} - {ex.StackTrace}", EventLogEntryType.Error);
                }
            }
            catch (Exception ex)
            {
                throw new JobExecutionException(ex);
            }
        }

        private void SendEmail(string fileAttachment, string ReportName)
        {
            try
            {
                string HostAdd = ConfigurationManager.AppSettings["Host"].ToString();
                string FromEmailAddress = ConfigurationManager.AppSettings["FromMail"].ToString();
                string ToEmailAddress = ConfigurationManager.AppSettings["ToMail"].ToString();
                string ccList = ConfigurationManager.AppSettings["CC_Mail"].ToString();
                string bccList = ConfigurationManager.AppSettings["BCC_Mail"].ToString();
                string Password = ConfigurationManager.AppSettings["Password"].ToString();

                string filePath = System.IO.Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory + "/Reports/MailContent.html");
                if (!string.IsNullOrEmpty(filePath))
                {
                    MailMessage mail = new MailMessage();
                    mail.From = new MailAddress(FromEmailAddress);
                    mail.Subject = string.Format("{0} - {1}", ReportName, DateTime.Now);
                    mail.IsBodyHtml = true;
                    mail.Body = File.ReadAllText(filePath);

                    string[] ToMuliId = ToEmailAddress.Split(',');
                    foreach (string ToEMailId in ToMuliId)
                    {
                        mail.To.Add(new MailAddress(ToEMailId)); //adding multiple TO Email Id  
                    }

                    if (!string.IsNullOrEmpty(ccList))
                    {
                        string[] CCId = ccList.Split(',');

                        foreach (string CCEmail in CCId)
                        {
                            mail.CC.Add(new MailAddress(CCEmail)); //Adding Multiple CC email Id  
                        }
                    }

                    if (!string.IsNullOrEmpty(bccList))
                    {
                        string[] bccid = bccList.Split(',');

                        foreach (string bccEmailId in bccid)
                        {
                            mail.Bcc.Add(new MailAddress(bccEmailId)); //Adding Multiple BCC email Id  
                        }
                    }

                    SmtpClient SmtpServer = new SmtpClient(HostAdd);
                    System.Net.Mail.Attachment attachment = new System.Net.Mail.Attachment(fileAttachment);
                    mail.Attachments.Add(attachment);
                    SmtpServer.Port = 587;
                    SmtpServer.UseDefaultCredentials = false;
                    SmtpServer.Credentials = new NetworkCredential(FromEmailAddress, Password);
                    SmtpServer.EnableSsl = true;
                    SmtpServer.Send(mail);
                }
                else
                {
                    EventLog.WriteEntry("Cron Email Service", $"Report File Path not found", EventLogEntryType.Error);
                }
            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("Cron Email Service", $"Exception occurred - {ex.Message} - {ex.StackTrace}", EventLogEntryType.Error);
            }
        }

        private string ExportToExcel(System.Data.DataTable dt, string label)
        {
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook wb = null;

            object misValue = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel.Worksheet ws = null;
            Microsoft.Office.Interop.Excel.Range rng = null;
            string filename = String.Empty;
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;

            try
            {
                filename = System.IO.Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory + "/Reports/" + label + "/" + string.Format("{0}.xlsx", label));

                wb = excel.Workbooks.Add(misValue);
                ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);

                for (int Idx = 0; Idx < dt.Columns.Count; Idx++)
                {
                    ws.Range["A1"].Offset[0, Idx].Value = dt.Columns[Idx].ColumnName;
                    ws.Range["A1"].Offset[0].EntireRow.Font.Bold = true;
                }

                if (label == "Stock Report")
                {
                    for (int Idx = 0; Idx < dt.Rows.Count; Idx++)
                    {
                        ws.Range["A2"].Offset[Idx].Resize[1, dt.Columns.Count].Value =
                        dt.Rows[Idx].ItemArray;
                    }
                }
                else if (label == "Sales Report")
                {
                    int f_idx = 0;
                    for (int Idx = 0; Idx < dt.Rows.Count; Idx++)
                    {
                        int g_idx = f_idx + (Idx == 0 ? Idx : 1);
                        var salesDataArray = dt.Rows[Idx].ItemArray;
                        salesDataArray[1] = "Given Below";
                        ws.Range["A2"].Offset[g_idx].Resize[1, dt.Columns.Count].Value = salesDataArray;
                        //dt.Rows[Idx].ItemArray;
                        if (dt.Rows[Idx]["Order Details"] != null)
                        {
                            DataTable dataTable2 = (DataTable)JsonConvert.DeserializeObject(dt.Rows[Idx]["Order Details"].ToString(), (typeof(DataTable)));
                            if (dataTable2 != null && dataTable2.Rows.Count > 0)
                            {
                                int A_cellrange = 3 + g_idx;
                                int D_cellrange = 3 + g_idx + dataTable2.Rows.Count;
                                for (int c_idx = 0; c_idx < dataTable2.Columns.Count; c_idx++)
                                {
                                    ws.Range["A3"].Offset[g_idx, c_idx + 1].Value = textInfo.ToTitleCase(dataTable2.Columns[c_idx].ColumnName.Replace("_", " "));
                                    ws.Range["A3"].Offset[g_idx].EntireRow.Font.Bold = true;
                                }
                                for (int idx = 0; idx < dataTable2.Rows.Count; idx++)
                                {
                                    f_idx = g_idx + idx;
                                    ws.Range["A4"].Offset[f_idx, 1].Resize[1, 4].Value = dataTable2.Rows[idx].ItemArray;
                                }
                                string _range = string.Format("A{0}:D{1}", A_cellrange, D_cellrange);
                                Excel.Range range = ws.Range[_range] as Excel.Range;
                                range.Rows.Group(misValue, misValue, misValue, misValue);
                                f_idx += 2;
                            }
                        }
                    }
                    ws.Outline.SummaryRow = XlSummaryRow.xlSummaryAbove;
                    ws.Outline.ShowLevels(1, 0);
                }

                ws.Columns.AutoFit();
                wb.RefreshAll();
                excel.Calculate();
                wb.SaveCopyAs(filename);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                wb.Close(true, misValue, misValue);
                excel.Quit();

                uint pid;
                HandleRef hwnd = new HandleRef(excel, (IntPtr)excel.Hwnd);
                GetWindowThreadProcessId((IntPtr)excel.Hwnd, out pid);
                KillProcess(pid, "EXCEL");

                Marshal.FinalReleaseComObject(wb);
                Marshal.FinalReleaseComObject(excel);
                excel = null;

            }
            catch (Exception ex)
            {
                EventLog.WriteEntry("Cron Email Service", $"Exception occurred - {ex.Message} - {ex.StackTrace} - {ex.InnerException}", EventLogEntryType.Error);
            }
            return filename;
        }
        
        private void KillProcess(uint pid, string processName)
        {
            // to kill current process of excel
            System.Diagnostics.Process[] AllProcesses = System.Diagnostics.Process.GetProcessesByName(processName);
            foreach (System.Diagnostics.Process process in AllProcesses)
            {
                if (process.Id == pid)
                {
                    process.Kill();
                }
            }
            AllProcesses = null;
        }
    }
}
