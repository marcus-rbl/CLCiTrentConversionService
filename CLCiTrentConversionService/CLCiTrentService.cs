using System;
using System.Collections.Generic;
using System.ServiceProcess;
using System.Threading;
using System.Text;
using System.IO;
using System.Net.Mail;
using Renci.SshNet;
using CsvHelper;
using System.Configuration;
using System.Security.Authentication;
using System.Net;
using System.Diagnostics;

namespace CLCiTrentConversionService
{
    public partial class CLCiTrentService : ServiceBase
    {
        Thread Thread;
        readonly AutoResetEvent StopEvent;
        string output = "";

        public CLCiTrentService()
        {
            InitializeComponent();

            StopEvent = new AutoResetEvent(initialState: false);

            eventLog1 = new System.Diagnostics.EventLog();
            if (!EventLog.SourceExists("CLCiTrent"))
            {
                EventLog.CreateEventSource(
                    "CLCiTrent", "CLCiTrentLog");
            }
            eventLog1.Source = "MySCLCiTrent";
            eventLog1.Log = "CLCiTrentLog";
        }

        protected override void OnStart(string[] args)
        {
            Thread = new Thread(ThreadStart);

            if (args.Length == 0)
            {
                Thread.Start(TimeSpan.Parse("04:00:00"));     //If time not specified in Start parameters use default of 4am               
            }
            else
            {
                Thread.Start(TimeSpan.Parse(args[0]));
            }
        }

        protected override void OnStop()
        {            
            if (!StopEvent.Set())
                Environment.FailFast("failed setting stop event");

            Thread.Join(1000);
        }

        void ThreadStart(object parameter)
        {
            while (!StopEvent.WaitOne(Timeout(timeOfDay: (TimeSpan)parameter)))
            {
                if (FetchFileBySftp())
                {
                    if (ProcessData())
                    {
                        CallWebService();
                    }                    
                }
            }
        }

        static TimeSpan Timeout(TimeSpan timeOfDay)
        {
            var timeout = timeOfDay - DateTime.Now.TimeOfDay;

            if (timeout < TimeSpan.Zero)
                timeout += TimeSpan.FromDays(1);

            return timeout;
        }

        private bool FetchFileBySftp()
        {
            string sftpHost = ConfigurationManager.AppSettings["sftpHost"];
            string sftpUsername = ConfigurationManager.AppSettings["sftpUsername"];
            string sftpPassword = ConfigurationManager.AppSettings["sftpPassword"];
            string remotePath = ConfigurationManager.AppSettings["remotePath"];
            string localPath = ConfigurationManager.AppSettings["localPath"];

            string fileName = DateTime.Now.ToString("yyyy-MM-dd") + ".csv";

            using (SftpClient sftp = new SftpClient(sftpHost, sftpUsername, sftpPassword))
            {
                try
                {
                    sftp.Connect();

                    if (sftp.Exists(remotePath + fileName))
                    {
                        using (Stream fileStream = File.OpenWrite(localPath + fileName))
                        {
                            sftp.DownloadFile(remotePath + fileName, fileStream);
                            eventLog1.WriteEntry(fileName + " has been donwloaded succesfully.");
                        }
                    }
                    else
                    {
                        eventLog1.WriteEntry("Today's file is not available: " + fileName);
                        SendHtmlEmail("Today's file is not available: " + fileName);
                        return false;
                    }
                    sftp.Disconnect();                    
                }
                catch (Exception exc)
                {
                    eventLog1.WriteEntry("Problem with sftp transfer:\n\n" + exc.ToString());
                    SendHtmlEmail("Problem with sftp transfer:<br/><br/>" + exc.ToString());
                    return false;
                }
            }
            return true;
        }

        private bool ProcessData()
        {
            try
            {
                eventLog1.WriteEntry("Processing Data");
                string outputPath = ConfigurationManager.AppSettings["outputPath"];
                string localPath = ConfigurationManager.AppSettings["localPath"];
                string fileName = DateTime.Now.ToString("yyyy-MM-dd") + ".csv";

                //Header
                output = "PER_REF_NO,TITLE,START_DATE,END_DATE,COMPLETED_I,COURSE_TYPE1,FAIL_I\n";

                var records = new List<OutputFile>();

                using (var reader = new StreamReader(localPath + fileName))

                using (var csv = new CsvReader(reader))
                {
                    csv.Read();
                    csv.ReadHeader();

                    while (csv.Read())
                    {
                        var record = new OutputFile
                        {
                            PER_REF_NO = csv.GetField("username"),
                            TITLE = csv.GetField("coursename"),
                            START_DATE = csv.GetField("timestarted"),
                            END_DATE = csv.GetField("timecompleted"),
                            COMPLETED_I = "T",
                            COURSE_TYPE1 = "On-Line",
                            FAIL_I = "F"
                        };

                        //convert unix time in seconds of original file to date with format of yyyyMMdd
                        DateTimeOffset dateTimeOffset = DateTimeOffset.FromUnixTimeSeconds(long.Parse(record.START_DATE));
                        DateTime dateTime = dateTimeOffset.UtcDateTime;
                        record.START_DATE = dateTime.ToString("yyyyMMdd");

                        dateTimeOffset = DateTimeOffset.FromUnixTimeSeconds(long.Parse(record.END_DATE));
                        dateTime = dateTimeOffset.UtcDateTime;
                        record.END_DATE = dateTime.ToString("yyyyMMdd");

                        //Only add the record if the Person Reference Number contains only numerical characters and is exactly 5 characters long (the only valid number)
                        if (int.TryParse(record.PER_REF_NO, out int n) && record.PER_REF_NO.Length == 5)
                        {
                            records.Add(record);
                            output += record.PER_REF_NO + "," + record.TITLE + "," + record.START_DATE + "," + record.END_DATE + "," + record.COMPLETED_I + "," + record.COURSE_TYPE1 + "," + record.FAIL_I + "\n";
                        }
                    }
                }

                eventLog1.WriteEntry("Writing output file.");

                using (var writer = new StreamWriter(outputPath + "itrent " + fileName))

                using (var csv = new CsvWriter(writer))
                {
                    csv.WriteRecords(records);
                }
                return true;
            }
            catch(Exception ex)
            {
                eventLog1.WriteEntry("Problem processing file " + ex.ToString());
                SendHtmlEmail("Problem processing file:<br/><br/> " + ex.ToString());
                return false;
            }
        }

        private bool SendHtmlEmail(string message)
        {            
            try
            {                
                var messageFrom = ConfigurationManager.AppSettings["messageFrom"];
                var bccMessageTo = ConfigurationManager.AppSettings["bccMessageTo"];
                var mailSubject = ConfigurationManager.AppSettings["messageSubject"];
                var messageTo = ConfigurationManager.AppSettings["messageTo"];

                var sb = new StringBuilder();
                sb.AppendLine("<html><body>");
                sb.AppendLine("<table style=\"width: 100%; text - align:left; border: none;\">");
                sb.AppendLine("<tr style=\"width:100%\">");
                sb.AppendLine("<td style=\"width:100%; font-family:Arial; font-size:medium 11px;\">");
                sb.AppendLine("</td></tr></table>");
                sb.AppendLine("</br>");
                sb.AppendLine("<table style=\"width: 100%; text - align:left; border: none;\">");
                sb.AppendLine("<tr style=\"width:100%\">");
                sb.AppendLine("<td style=\"width:100%; font-family:Arial; font-size:medium 11px;\">");
                sb.AppendLine(message);
                sb.AppendLine("</td></tr></table>");
                sb.AppendLine("<br/>");
                sb.AppendLine("<table style=\"width: 100%; text - align:left; border: none;\">");
                sb.AppendLine("<tr style=\"width:100%\">");
                sb.AppendLine("<td style=\"width:100%; font-family:Arial; font-size:medium 11px; font-weight:bold;\">");
                sb.AppendLine("IM&T Solutions Delivery Team");
                sb.AppendLine("</td></tr></table>");
                sb.AppendLine("</html></body>");

                //send email
                var emailSent = SmtpMailSender.SendMailmessage(sb.ToString(), mailSubject, messageTo, messageFrom, bccMessageTo);
                return emailSent;
            }
            catch (Exception ex)
            {
                eventLog1.WriteEntry("Problem sending email: " + ex.ToString());
                return false;
            }
        }

        public class SmtpMailSender
        {
            public static bool SendMailmessage(string htmlData, string mailSubject, string messageTo, string messageFrom, string bccMessageTo)
            {
                try
                {
                    var client = new SmtpClient();
                    client.Port = 25;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = "smtp.britishlegion.org.uk";

                    var mail = new MailMessage();
                    mail.From = new MailAddress(messageFrom);
                    mail.To.Add(new MailAddress(messageTo));
                    mail.Bcc.Add(new MailAddress(bccMessageTo));
                    mail.Subject = mailSubject;
                    mail.IsBodyHtml = true;
                    mail.Body = htmlData;
                    client.Send(mail);
                    return true;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        private void CallWebService()
        {
            string p_conv_type = "LEARNEVENTS";
            string p_conv_dir = "";
            string p_conv_file = output;
            string p_people_id = "PERREF";
            string p_fs = ",";
            string p_org_name = "The Royal British Legion";
            string p_user_nm = "TRBLWEBSERVICES";
            string p_user_pwd = "FAJz?YJzejD2Y746";

            string p_error_msg = "";
            string p_exc_file = "";
            string p_log_file = "";
            string p_suc_file = "";
            string p_queue_id = "";

            com.webitrent.ce0286de.ETADM086SSService service = new com.webitrent.ce0286de.ETADM086SSService();

            try
            {
                const SslProtocols _Tls12 = (SslProtocols)0x00000C00;
                const SecurityProtocolType Tls12 = (SecurityProtocolType)_Tls12;
                ServicePointManager.SecurityProtocol = Tls12;

                int i = service.RUN_CONV_NEW(p_conv_type, p_conv_dir, p_conv_file, p_people_id, p_fs, p_org_name, p_user_nm, p_user_pwd, out p_log_file, out p_exc_file, out p_suc_file, out p_queue_id, out p_error_msg);

                string message = "CLC to iTrent upload is complete.<br/><br/>" +
                                    "p_error_msg: " + p_error_msg + "<br/><br/>" +
                                    "p_exc_file: " + p_exc_file + "<br/><br/>" +
                                    "p_log_file: " + p_log_file + "<br/><br/>" +
                                    "p_suc_file: " + p_suc_file + "<br/><br/>" +
                                    "p_queue_id: " + p_queue_id;
                SendHtmlEmail(message);

            }
            catch (Exception exc)
            {
                eventLog1.WriteEntry("Calling Web Service failed: " + exc.ToString());
                SendHtmlEmail("Calling Web Service failed:<br/><br/> " + exc.ToString());
            }
        }

        private class OutputFile
        {
            public string PER_REF_NO { get; set; }
            public string TITLE { get; set; }
            public string START_DATE { get; set; }
            public string END_DATE { get; set; }
            public string COMPLETED_I { get; set; }
            public string COURSE_TYPE1 { get; set; }
            public string FAIL_I { get; set; }
        }
    }
}
