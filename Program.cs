using System;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SendMailFromOutlook
{
    class Program
    {
        const string attachmentPath = @"C:\Users\tsasajima\Desktop\BTMU\Office 365 Presentation (wide).pptx";
        const string sender = "testuser1@test0148.onmicrosoft.com";
        //static string[] recipients = new string[] { "testuser4@test0148.onmicrosoft.com" };
        //static string[] recipients = new string[] { "testuser5@test0148.onmicrosoft.com", "testuser6@test0148.onmicrosoft.com" };
        //static string[] recipients = new string[] { "testuser7@test0148.onmicrosoft.com", "testuser8@test0148.onmicrosoft.com", "testuser9@test0148.onmicrosoft.com" };
        static string[] recipients = new string[] { "testuser10@test0148.onmicrosoft.com", "testuser11@test0148.onmicrosoft.com" };
        const string subjectFormat = "Discovery Search Mail Item {0}";
        const string discoveryString = "あ";
        const int totalMailCount = 2000;
        const int discoveryCount = totalMailCount / 20;
        const int attachmentCount = totalMailCount / 4;
        const int bodyLength = 1024;
        const int lineLength = 64;
        const int progressMessageInterval = 20;
        const int subjectNumberOffset = 2000;
        static Outlook.Application outlook = null;
        static Random random = new Random();

        static void Main(string[] args)
        {
            Console.WriteLine("Hit enter key to start.");
            while (Console.ReadKey().Key != ConsoleKey.Enter) { }
            outlook = new Outlook.Application();
            Outlook.NameSpace nameSpace = outlook.GetNamespace("MAPI");
            try {
                nameSpace.Logon("", "", Missing.Value, Missing.Value);
                for (int i = 1; i <= totalMailCount; i++) {
                    Outlook.MailItem mail = CreateMail(i, null, recipients[i % recipients.Length]);
                    int mod = totalMailCount / attachmentCount;
                    if (mod > 0 && i % mod == 0) {
                        mail.Attachments.Add(attachmentPath, Outlook.OlAttachmentType.olByValue, 1, attachmentPath);
                    }
                    mail.Send();
                    if (i % progressMessageInterval == 0) {
                        Console.WriteLine(string.Format("{0} messages queued in outbox.", i));
                    }
                }
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
                if (ex.InnerException != null) {
                    Console.WriteLine(ex.InnerException.Message);
                }
                Console.WriteLine(ex.StackTrace);
            } finally {
                Console.WriteLine("Queuing finished.");
                Outlook.Folder outbox = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox) as Outlook.Folder;
                while (true) {
                    int itemCount = outbox.Items.Count;
                    if (itemCount == 0) {
                        break;
                    }
                    Console.WriteLine("{0} messages left in outbox. Wait until all messages flushed from outbox.", itemCount);
                    Thread.Sleep(10000);
                }
                Console.WriteLine("All messages has been sent.");
                Console.WriteLine("Hit enter key to quit.");
                while (Console.ReadKey().Key != ConsoleKey.Enter) { }
                nameSpace.Logoff();
                outlook.Quit();
            }
        }

        static Outlook.MailItem CreateMail(int num, string sender, string recipient)
        {
            var mail = outlook.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
            mail.Subject = string.Format(subjectFormat, num + subjectNumberOffset);
            if (!string.IsNullOrEmpty(sender)) {
                mail.Sender = outlook.Session.CreateRecipient(sender).AddressEntry;
            }
            mail.Body = CreateRandomString(num);
            mail.Recipients.Add(recipient);
            return mail;
        }

        static string CreateRandomString(int num)
        {
            const string pool = "abcdefghijklmnopqrstuvwxyz0123456789";
            var chars = Enumerable.Range(0, bodyLength).Select(x => pool[random.Next(0, pool.Length)]).ToList();
            string s = new string(chars.ToArray());
            int mod = totalMailCount / discoveryCount;
            if (mod > 0 && num % mod == 0) {
                s = s.Insert(random.Next(0, bodyLength - 1), discoveryString);
            }
            return Regex.Replace(s, "(.{" + lineLength + "})", "$1" + Environment.NewLine);
        }
    }
}
