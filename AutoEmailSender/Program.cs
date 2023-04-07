using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Collections.Generic;
using System.Configuration;
using System.Threading;
using System.Threading.Tasks;

namespace AutoEmailSender
{
    internal class Program
    {
        static readonly string networkSharePath = ConfigurationManager.AppSettings["NetworkSharePath"];
        static readonly string fileExtension = ConfigurationManager.AppSettings["FileExtension"];
        static readonly string fileNamePrefix = ConfigurationManager.AppSettings["FileNamePrefix"];
        static readonly string fromEmail = ConfigurationManager.AppSettings["FromEmail"];
        static readonly string fromEmailPassword = ConfigurationManager.AppSettings["FromEmailPassword"];
        static readonly string toEmail = ConfigurationManager.AppSettings["ToEmail"];
        static readonly string reminderToEmail = ConfigurationManager.AppSettings["ReminderToEmail"];
        static readonly string smtpServer = ConfigurationManager.AppSettings["SmtpServer"];
        static readonly int smtpPort = Convert.ToInt32(ConfigurationManager.AppSettings["SmtpPort"]);
        static readonly bool enableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSsl"]);
        static bool enbaleSend = false;

        static async Task Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Title = "FMD邮件自动推送";
            var task1 = ScheduleTask(new TimeSpan(14, 10, 0), TaskAt14pm, "TaskAt14pm");
            var task2 = ScheduleTask(new TimeSpan(16, 0, 0), TaskAt16pm, "TaskAt16pm");
            Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss}  FMD邮件自动发送任务已启用");
            await Task.WhenAll(task1, task2);
        }

        private static async Task ScheduleTask(TimeSpan targetTime, Action task, string taskName)
        {
            while (true)
            {
                DateTime now = DateTime.Now;
                DateTime todayTargetTime = DateTime.Today.Add(targetTime);
                DateTime nextOccurrence = todayTargetTime > now ? todayTargetTime : todayTargetTime.AddDays(1);
                TimeSpan timeUntilNextTask = nextOccurrence - now;
                await Task.Delay(timeUntilNextTask);
                task();
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} Task executed: {taskName} 已执行");
            }
        }

        private static void TaskAt14pm()
        {

            DateTime now = DateTime.Now;
            DayOfWeek dayOfWeek = now.DayOfWeek;

            if (dayOfWeek == DayOfWeek.Saturday || dayOfWeek == DayOfWeek.Sunday)
            {
                
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} 今天是周六或周日，不会发送邮件");
            }
            else
            {
                var filesToSend = GetFilesFromNetworkShare(networkSharePath, fileExtension, fileNamePrefix);
                if (filesToSend.Count != 0)
                {
                    SendEmailWithAttachments(fromEmail, fromEmailPassword, toEmail, smtpServer, smtpPort, enableSsl, filesToSend);
                    enbaleSend = true; 
                }
                else
                {
                    SendReminderEmail(fromEmail, fromEmailPassword, reminderToEmail, smtpServer, smtpPort, enableSsl);
                    
                }
            }            
        }

        private static void TaskAt16pm()
        {

            DateTime now = DateTime.Now;
            DayOfWeek dayOfWeek = now.DayOfWeek;

            if (dayOfWeek == DayOfWeek.Saturday || dayOfWeek == DayOfWeek.Sunday)
            {
                
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} 今天是周六或周日，不会发送邮件");
            }
            
            else
            {
                var filesToSend = GetFilesFromNetworkShare(networkSharePath, fileExtension, fileNamePrefix);
                if (filesToSend.Count != 0 && enbaleSend == false)
                {
                    SendEmailWithAttachments(fromEmail, fromEmailPassword, toEmail, smtpServer, smtpPort, enableSsl, filesToSend);                  

                }
                else if (filesToSend.Count != 0 && enbaleSend == true)
                {
                    Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} 已执行过14:00 邮件推送任务，无需重复发送");
                    enbaleSend = false;
                }
                else
                {
                    // 保存原始控制台颜色
                    ConsoleColor originalColor = Console.ForegroundColor;
                    // 将颜色更改为黄色并打印这一行
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} 没有找到文件");
                    // 将控制台颜色重置为原始颜色
                    Console.ForegroundColor = originalColor;
                }
            }
            
        }
        
        private static List<string> GetFilesFromNetworkShare(string networkSharePath, string fileExtension, string fileNamePrefix)
        {
            List<string> filesToSend = new List<string>();
            try
            {
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} 正在从网络共享路径 '{networkSharePath}' 获取以'{fileNamePrefix}'开头且创建于当天的{fileExtension}文件...");
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
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} 找到 {filesToSend.Count} 个符合条件的{fileExtension}文件");
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
                // 保存原始控制台颜色
                ConsoleColor originalColor = Console.ForegroundColor;

                // 将颜色更改为黄色并打印这一行
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} 没有找到符合条件的文件，不会发送邮件");
                // 将控制台颜色重置为原始颜色
                Console.ForegroundColor = originalColor;
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
                        Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} FMD邮件已成功发送！");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发送邮件时出错: {ex.Message}");
            }
        }
        
        private static void SendReminderEmail(string fromEmail, string fromEmailPassword, string toEmail, string smtpServer, int smtpPort, bool enableSsl)
        {
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(fromEmail);
                    mail.To.Add(toEmail);
                    mail.Subject = "FMD提醒";
                    mail.Body = $"今天下午{DateTime.Now:yyyy-MM-dd HH:mm:ss} 没有找到符合条件的FMD文件，请注意";
                    using (SmtpClient smtp = new SmtpClient(smtpServer, smtpPort)) // 请根据你的邮箱服务提供商设置SMTP服务器地址和端口
                    {
                        smtp.Credentials = new NetworkCredential(fromEmail, fromEmailPassword);
                        smtp.EnableSsl = enableSsl; // 如果需要，请启用SSL
                        smtp.Send(mail);
                        // 保存原始控制台颜色
                        ConsoleColor originalColor = Console.ForegroundColor;

                        // 将颜色更改为黄色并打印这一行
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} FMD提醒邮件已成功发送！");
                        Console.ForegroundColor = originalColor;
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
