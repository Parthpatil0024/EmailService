using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace SampleService
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 180000; //number in miliseconds  
            timer.Enabled = true;
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {

            WriteToFile("Service recalled at " + DateTime.Now);

            try
            {
               
                DataTable dt = new DataTable();
                string query = "SELECT top 1 * FROM tblEmail WHERE IsEmailSent=0";
                string constr = ConfigurationManager.ConnectionStrings["DBTContext"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;

                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                        }
                    }
                }
               
                foreach (DataRow row in dt.Rows)
                {
                    string ToEmail = row["ToEmail"].ToString();
                    
                    string Subject = row["Subject"].ToString();
                    string Body = row["Body"].ToString();
                    int Id = Convert.ToInt32(row["Id"]);
                    bool res= SendEmail(ToEmail, Subject, Body);
                    if (res)
                    {
                        UpdateStatus(Id);
                    }

                } }
            catch (Exception ex)
            {
                WriteToFile("Data Fetch Failed " + ex.InnerException.Message);
            }


                }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }


        public bool SendEmail(string ToEmail,string Subject,string Body)
        {
            bool result=false;
            try
            {
                
                var senderEmail = new MailAddress("prp9096@gmail.com", "ShopNow");
                var receiverEmail = new MailAddress(ToEmail, "Receiver");
                var password = "ggarecaijgnpfgea";
                var sub = Subject;
                var body = Body;

                MailMessage message = new MailMessage();
                message.To.Add(ToEmail);// Email-ID of Receiver  
                message.Subject = sub;// Subject of Email  
                message.From = senderEmail;// Email-ID of Sender  
                message.IsBodyHtml = true;


                message.Body = body;
                SmtpClient SmtpMail = new SmtpClient();

                SmtpMail.Host = "smtp.gmail.com";
                SmtpMail.Port = 587;
                SmtpMail.EnableSsl = true;
             
                SmtpMail.DeliveryMethod = SmtpDeliveryMethod.Network;
                SmtpMail.UseDefaultCredentials = false;
                SmtpMail.Credentials = new NetworkCredential(senderEmail.Address, password);
                
                SmtpMail.Send(message);
                
                WriteToFile("Email sent  for" + ToEmail);
                result = true;



            }
            catch (Exception ex)
            {
                WriteToFile("Email sending failed for " + ToEmail + "due to " + ex.InnerException.Message);
                
            }

            return result;
        }

        public void UpdateStatus(int  Id)
        {
            
            string conStr = ConfigurationManager.ConnectionStrings["DBTContext"].ConnectionString;

            
            string query = "update  tblEmail set IsEmailSent= '" + 1 +
                 
                "' where Id =" + Id;



            try
            {
                using (SqlConnection conn = new SqlConnection(conStr)) 
                { 
                 SqlCommand comm = new SqlCommand(query, conn);

                    comm.ExecuteNonQuery();

                    WriteToFile("Status Updated");

                }
            }
            catch (Exception ex)
            {
                WriteToFile("Error Updating Status "+ ex.InnerException.Message);
            }
            
           
        }



    }
}
