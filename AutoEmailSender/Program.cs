using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using System.Configuration;
using System.Threading;

namespace AutoEmailSender
{
    internal class Program
    {
        static string networkSharePath = ConfigurationManager.AppSettings["NetworkSharePath"];
        static string fileExtension = ConfigurationManager.AppSettings["FileExtension"];
        static string fileNamePrefix = ConfigurationManager.AppSettings["FileNamePrefix"];
        static string fromEmail = ConfigurationManager.AppSettings["FromEmail"];
        static string fromEmailPassword = ConfigurationManager.AppSettings["FromEmailPassword"];
        static string toEmail = ConfigurationManager.AppSettings["ToEmail"];
        static string reminderToEmail = ConfigurationManager.AppSettings["ReminderToEmail"];
        static string smtpServer = ConfigurationManager.AppSettings["SmtpServer"];
        static int smtpPort = Convert.ToInt32(ConfigurationManager.AppSettings["SmtpPort"]);
        static int startTime = Convert.ToInt32(ConfigurationManager.AppSettings["StartTime"]);
        static bool enableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSsl"]);
        static bool enbaleSend = false;
        static void Main(string[] args)
       
        {
            
            TimeSpan timeUntilNext14pm = GetNextOccurrence(new TimeSpan(14, 0, 0));
            TimeSpan timeUntilNext16pm = GetNextOccurrence(new TimeSpan(16, 0, 0));
            SetTimerForNextTask(timeUntilNext14pm, TaskAt14pm);
            SetTimerForNextTask(timeUntilNext16pm, TaskAt16pm);
     
            Console.WriteLine("设置多个任务的启动时间完毕");
            Console.WriteLine("按Enter键结束...");
            Console.ReadLine();
        }
        
        private static void TaskAt14pm()
        {

            DateTime now = DateTime.Now;
            DayOfWeek dayOfWeek = now.DayOfWeek;

            if (dayOfWeek == DayOfWeek.Saturday || dayOfWeek == DayOfWeek.Sunday)
            {
                Console.WriteLine("今天是周六或周日，不会发送邮件。");
            }
            else
            {
                var filesToSend = GetFilesFromNetworkShare(networkSharePath, fileExtension, fileNamePrefix);
                if (filesToSend.Count != 0)
                {
                    SendEmailWithAttachments(fromEmail, fromEmailPassword, toEmail, smtpServer, smtpPort, enableSsl, filesToSend);
                    Console.WriteLine("已执行14:00 PM邮件推送任务。当前时间：{0:HH:mm:ss}", DateTime.Now);
                    enbaleSend = true; 

                }
                else
                {
                    SendReminderEmail(fromEmail, fromEmailPassword, reminderToEmail, smtpServer, smtpPort, enableSsl, startTime);
                    Console.WriteLine("没有找到文件。当前时间：{0:HH:mm:ss}", DateTime.Now);
                }
            }
            
        }

        private static void TaskAt16pm()
        {

            DateTime now = DateTime.Now;
            DayOfWeek dayOfWeek = now.DayOfWeek;

            if (dayOfWeek == DayOfWeek.Saturday || dayOfWeek == DayOfWeek.Sunday)
            {
                Console.WriteLine("今天是周六或周日，不会发送邮件。");
            }
            else
            {
                var filesToSend = GetFilesFromNetworkShare(networkSharePath, fileExtension, fileNamePrefix);
                if (filesToSend.Count != 0 && enbaleSend == false)
                {
                    SendEmailWithAttachments(fromEmail, fromEmailPassword, toEmail, smtpServer, smtpPort, enableSsl, filesToSend);
                    Console.WriteLine("已执行16:00 PM邮件推送任务。当前时间：{0:HH:mm:ss}", DateTime.Now);

                }
                else if (filesToSend.Count != 0 && enbaleSend == true)
                {
                    Console.WriteLine("已执行过14:00 PM邮件推送任务，无需重复发送。当前时间：{0:HH:mm:ss}", DateTime.Now);
                    enbaleSend = false;
                }
                else
                {
                    Console.WriteLine("没有找到文件。当前时间：{0:HH:mm:ss}", DateTime.Now);
                }
            }
            
        }
        
        private static void SetTimerForNextTask(TimeSpan timeUntilNextTask, Action task)
        {
            // 创建定时器，并设置触发时间
            Timer timer = new Timer(OnTimedEvent, task, timeUntilNextTask, Timeout.InfiniteTimeSpan);
        }

        private static void OnTimedEvent(object state)
        {
            // 获取要执行的任务
            Action task = (Action)state;

            // 执行任务
            task();

            // 重新设置定时器以在明天的相同时间执行任务
            TimeSpan timeUntilNextOccurrence = GetNextOccurrence(TimeSpan.FromDays(1));
            SetTimerForNextTask(timeUntilNextOccurrence, task);
        }

        private static TimeSpan GetNextOccurrence(TimeSpan targetTime)
        {
            DateTime now = DateTime.Now;
            DateTime todayTargetTime = DateTime.Today.Add(targetTime);
            DateTime nextOccurrence = todayTargetTime > now ? todayTargetTime : todayTargetTime.AddDays(1);//如果已经过了指定时间，就等于明天定时的小时时间
            return nextOccurrence - now;
        }
        
        private static List<string> GetFilesFromNetworkShare(string networkSharePath, string fileExtension, string fileNamePrefix)
        {
            List<string> filesToSend = new List<string>();
            try
            {
                Console.WriteLine($"正在从网络共享路径 '{networkSharePath}' 获取以'{fileNamePrefix}'开头且创建于当天的{fileExtension}文件...");
                string[] files = Directory.GetFiles(networkSharePath, "*" + fileExtension);
                DateTime today = DateTime.Today;

                foreach (string file in files)
                {
                    FileInfo fileInfo = new FileInfo(file);

                    if (fileInfo.Name.StartsWith(fileNamePrefix) && fileInfo.CreationTime.Date == today)
                    {
                        filesToSend.Add(file);
                    }
                }
                Console.WriteLine($"找到 {filesToSend.Count} 个符合条件的{fileExtension}文件:");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取文件时出错: {ex.Message}");
            }

            return filesToSend;
        }

        private static void SendEmailWithAttachments(string fromEmail, string fromEmailPassword, string toEmail, string smtpServer, int smtpPort, bool enableSsl, List<string> filesToSend)
        {
            if (filesToSend.Count == 0)
            {
                Console.WriteLine("没有找到符合条件的文件，不会发送邮件。");
                return;
            }
            DateTime today = DateTime.Today;
            DayOfWeek dayOfWeek = today.DayOfWeek;
            if (dayOfWeek == DayOfWeek.Sunday)
            {
                Console.WriteLine("今天是周六或周日，不会发送邮件。");
                return;
            }
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(fromEmail);
                    mail.To.Add(toEmail);
                    mail.Subject = "FMD PDF report";
                    mail.Body = $"Please find attached {filesToSend.Count} PDF files.";

                    foreach (string file in filesToSend)
                    {
                        mail.Attachments.Add(new Attachment(file));
                    }
                    using (SmtpClient smtp = new SmtpClient(smtpServer, smtpPort)) // 请根据你的邮箱服务提供商设置SMTP服务器地址和端口
                    {
                        smtp.Credentials = new NetworkCredential(fromEmail, fromEmailPassword);
                        smtp.EnableSsl = enableSsl; // 如果需要，请启用SSL
                        smtp.Send(mail);
                        Console.WriteLine("邮件已成功发送！");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发送邮件时出错: {ex.Message}");
            }
        }
        
        private static void SendReminderEmail(string fromEmail, string fromEmailPassword, string toEmail, string smtpServer, int smtpPort, bool enableSsl, int startTime)
        {
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(fromEmail);
                    mail.To.Add(toEmail);
                    mail.Subject = "FMD提醒";
                    mail.Body = "今天下午" + startTime.ToString() + "点没有找到符合条件的FMD文件，请注意。";
                    using (SmtpClient smtp = new SmtpClient(smtpServer, smtpPort)) // 请根据你的邮箱服务提供商设置SMTP服务器地址和端口
                    {
                        smtp.Credentials = new NetworkCredential(fromEmail, fromEmailPassword);
                        smtp.EnableSsl = enableSsl; // 如果需要，请启用SSL
                        smtp.Send(mail);

                        Console.WriteLine("提醒邮件已成功发送！");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发送提醒邮件时出错: {ex.Message}");
            }
        }
    }
}
