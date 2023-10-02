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

namespace CronService_Processor
{
    public class ScheduledJobs : IJob
    {
        public void Execute(IJobExecutionContext context)
        {
            try
            {
                SendEmail("Sathishkumarr21@gmail.com", null, null, "Sample", "Sample Mail");
            }
            catch (Exception ex)
            {
                throw new JobExecutionException(ex);
            }
        }

        public void SendEmail(String ToEmail, string cc, string bcc, String Subj, string Message)
        {
            //Reading sender Email credential from web.config file  
            try
            {
                string HostAdd = ConfigurationManager.AppSettings["Host"].ToString();
                string FromEmailid = ConfigurationManager.AppSettings["FromMail"].ToString();
                string Pass = ConfigurationManager.AppSettings["Password"].ToString();

                //creating the object of MailMessage  
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(FromEmailid); //From Email Id  
                mailMessage.Subject = Subj; //Subject of Email  
                mailMessage.Body = Message; //body or message of Email  
                mailMessage.IsBodyHtml = true;

                string[] ToMuliId = ToEmail.Split(',');
                foreach (string ToEMailId in ToMuliId)
                {
                    mailMessage.To.Add(new MailAddress(ToEMailId)); //adding multiple TO Email Id  
                }

                if (!string.IsNullOrEmpty(cc))
                {
                    string[] CCId = cc.Split(',');

                    foreach (string CCEmail in CCId)
                    {
                        mailMessage.CC.Add(new MailAddress(CCEmail)); //Adding Multiple CC email Id  
                    }
                }

                if (!string.IsNullOrEmpty(bcc))
                {
                    string[] bccid = bcc.Split(',');

                    foreach (string bccEmailId in bccid)
                    {
                        mailMessage.Bcc.Add(new MailAddress(bccEmailId)); //Adding Multiple BCC email Id  
                    }
                }

                SmtpClient smtp = new SmtpClient(HostAdd);  // creating object of smptpclient  
                smtp.UseDefaultCredentials = false;
                NetworkCredential NetworkCred = new NetworkCredential();
                NetworkCred.UserName = mailMessage.From.Address;
                NetworkCred.Password = Pass;
                smtp.Credentials = NetworkCred;
                smtp.EnableSsl = true;
                smtp.Port = 587;
                smtp.Send(mailMessage); //sending Email 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
