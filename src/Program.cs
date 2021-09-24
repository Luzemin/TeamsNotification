using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Timers;

namespace ConsoleApp
{
    public class Program
    {
        private static readonly string templatePath = @"D:\Template.png";
        private static readonly string teamsIconSavePath = @"D:\Teams.png";
        private static readonly int captureX = 1670;
        private static readonly int captureY = 1045;
        private static readonly int timerIntervalMinutes = 3;
        private static readonly string emailTo = "your.email@qq.com";

        static void Main(string[] args)
        {
            Log($"start");
            SetTimer();
            Console.ReadLine();
        }

        private static void SetTimer()
        {
            Timer aTimer = new Timer(timerIntervalMinutes * 60 * 1000);
            aTimer.Elapsed += OnTimedEvent;
            aTimer.AutoReset = true;
            aTimer.Enabled = true;
        }

        private static void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            Log($"time to check");
            Log($"captrue Teams status");
            var teams = CaptrueTeamsIcon();
            var template = new Bitmap(Image.FromFile(templatePath));
            Log($"compare with template image");
            if (IsSamePic(teams, template))
            {
                Log($"teams status has changed, will send email");
                SendMail("Teams Message Notification", "Your Teams status has changed.", emailTo);
                Log($"email sent complete");
            }
            else
            {
                Log($"teams status has not changed");
            }
        }

        private static Bitmap CaptrueTeamsIcon(bool boolSave = false)
        {
            var bit = new Bitmap(30, 30);
            Graphics g = Graphics.FromImage(bit);
            g.CopyFromScreen(captureX, captureY, 0, 0, bit.Size);
            if (boolSave)
            {
                bit.Save(teamsIconSavePath, ImageFormat.Png);
            }
            g.Dispose();
            return bit;
        }

        private static bool IsSamePic(Bitmap firstImage, Bitmap secondImage)
        {
            MemoryStream ms = new MemoryStream();
            firstImage.Save(ms, ImageFormat.Png);
            string firstBitmap = Convert.ToBase64String(ms.ToArray());
            ms.Position = 0;
            secondImage.Save(ms, ImageFormat.Png);
            string secondBitmap = Convert.ToBase64String(ms.ToArray());
            ms.Dispose();
            return firstBitmap.Equals(secondBitmap);
        }

        public static bool SendMail(string subject, string body, string mailTo)
        {
            //email host config
            string emailSender = "your.email@qq.com";
            string emailPassword = "your.email.auth.code";
            string emailHost = "smtp.qq.com";
            int emailPort = 25;

            using (MailMessage message = new MailMessage())
            {
                message.From = new MailAddress(emailSender, "Teams Alarm", System.Text.Encoding.UTF8);
                message.To.Add(mailTo);
                message.Subject = subject;
                message.IsBodyHtml = true;
                message.BodyEncoding = System.Text.Encoding.UTF8;
                message.Body = body;
                using (SmtpClient client = new SmtpClient())
                {
                    client.Port = emailPort;
                    client.Host = emailHost;
                    client.Credentials = new NetworkCredential(emailSender, emailPassword);
                    try
                    {
                        client.Send(message);
                        Log($"email sent success");
                    }
                    catch (Exception ex)
                    {
                        Log($"email sent failed:{ex.Message}");
                        return false;
                    }
                }
            }
            return true;
        }

        public static void Log(string msg)
        {
            Console.WriteLine($"[{DateTime.Now}] {msg}");
        }
    }
}

