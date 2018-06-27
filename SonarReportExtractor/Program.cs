using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SonarReportExtractor
{

    public class Program
    {
        private static void RunSonarQube()
        {
            // System.Diagnostics.Process.Start("E:\\Source\\Sonar\\sonarqube-5.6.3\\bin\\windows-x86-64\\StartSonar.bat");
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
            // startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            //startInfo.FileName = "E:\\Source\\Sonar\\sonarqube-5.6.3\\bin\\windows-x86-64\\StartSonar.bat";
            startInfo.FileName = "StartSonar.bat";
            //startInfo.Arguments = "cd E:\\Source\\Sonar\\sonarqube-5.6.3\\bin\\windows-x86-64\\StartSonar.bat";
            process.StartInfo = startInfo;
            //process.StartInfo.v
            process.Start();
        }

        private static void SendEmailUsingSMTP(string attachmentFilename)
        {
            string sonarUrl = ConfigurationManager.AppSettings["sonarUrl"];

            string mailFrom = ConfigurationManager.AppSettings["mailFrom"];
            string mailTo = ConfigurationManager.AppSettings["mailTo"];
            
            try
            {
                string mailPassword = DecryptString(ConfigurationManager.AppSettings["password"]);
                Log("Preparing Email..\n");

               NetworkCredential mailCredential = new NetworkCredential(mailFrom, mailPassword);

                MailMessage mail = new MailMessage(mailFrom, mailTo);
                SmtpClient client = new SmtpClient();
                client.Port = Convert.ToInt32(ConfigurationManager.AppSettings["smtpPort"]);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
              //  client.UseDefaultCredentials = false;
                client.EnableSsl = true;

                client.Credentials = mailCredential;
                client.Host = ConfigurationManager.AppSettings["smtpHost"];
                var dateFormat = ConfigurationManager.AppSettings["DateFormat"].ToString();

                mail.Subject = ConfigurationManager.AppSettings["mailSubject"].Replace("REPORTDATE", DateTime.Now.ToString(dateFormat));

                mail.Body = ConfigurationManager.AppSettings["mailBody"].ToString().Replace("REPORTDATE", DateTime.Now.ToString(dateFormat)).Replace("SONARURL", sonarUrl);

                // mail.Body = string.Format($"Hi Team,<br><br>{ConfigurationManager.AppSettings["mailBody"]} { DateTime.Now.ToString("MMM yyyy") },<br> Sonar Url :<br>{sonarUrl}");

                mail.IsBodyHtml = true;

                if (attachmentFilename != null)
                {
                    Attachment attachment = new Attachment(attachmentFilename, MediaTypeNames.Application.Octet);
                    attachment.Name = string.Format("{0}{1}", ConfigurationManager.AppSettings["attachmentName"], DateTime.Now.ToString("yyyyMMdd"));
                    ContentDisposition disposition = attachment.ContentDisposition;
                    disposition.CreationDate = File.GetCreationTime(attachmentFilename);
                    disposition.ModificationDate = File.GetLastWriteTime(attachmentFilename);
                    disposition.ReadDate = File.GetLastAccessTime(attachmentFilename);
                    disposition.FileName = Path.GetFileName(attachmentFilename);
                    disposition.Size = new FileInfo(attachmentFilename).Length;
                    disposition.DispositionType = DispositionTypeNames.Attachment;
                    mail.Attachments.Add(attachment);
                }

                client.Send(mail);
                Log($"Email sent successfully to {mailTo}");
               
            }
            catch
            {
                Log($"Email sending failed \n From :{mailFrom} To :{mailTo}");
                throw;
            }
        }

        public static void SendEmailUsingOUTLOOK(string attachmentName)
        {
            string sonarUrl = ConfigurationManager.AppSettings["sonarUrl"];

            string mailFrom = ConfigurationManager.AppSettings["mailFrom"];
            string mailTo = ConfigurationManager.AppSettings["mailTo"];
            try
            {
                Log("Preparing mail..\n");
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody.
                //add the body of the email
                var dateFormat = ConfigurationManager.AppSettings["DateFormat"].ToString();
                oMsg.HTMLBody = ConfigurationManager.AppSettings["mailBody"].ToString().Replace("REPORTDATE", DateTime.Now.ToString(dateFormat)).Replace("SONARURL", sonarUrl);
                //"Hi Team,<br>" + ConfigurationManager.AppSettings["mailSubject"] + DateTime.Now.ToString("MMM yyyy");
                //Add an attachment.
               // String sDisplayName = string.Format("{0}", attachmentName);
                int iPosition = (int)oMsg.Body.Length + 1;
                int iAttachType = (int)Outlook.OlAttachmentType.olByValue;
                //now attached the file
                Outlook.Attachment oAttach = oMsg.Attachments.Add
                                             (attachmentName, iAttachType, iPosition);

                //Subject line
                oMsg.Subject = ConfigurationManager.AppSettings["mailSubject"].Replace("REPORTDATE", DateTime.Now.ToString(dateFormat));
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add(mailTo);
                oRecip.Resolve();
                // Send.
                oMsg.Send();
                // Clean up.
                oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;

                Log(string.Format("Email sent sucessfully to {0}", mailTo));
               

            }//end of try block
            catch
            {
                Log($"Email sending failed \n From :{mailFrom} To :{mailTo}");
                throw;
            }//end of catch
        }//end of Em

        private static void Log(string msg)
        {
            var date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Console.WriteLine(string.Format("[{0}] - {1}", date, msg));
        }
        private static void Main(string[] args)
        {
           
            try
            {
                string url = ConfigurationManager.AppSettings["sonardashboardUrl"];
                string filePath = ConfigurationManager.AppSettings["screenCaptureSaveLocation"];
                filePath = filePath + DateTime.Now.ToString("yyyyMMdd") + "." + "pdf";
                Log("Verifying Sonarqube Availability.");

                CheckSonarAvailability();
                Log("Sonarqube is running....");
                CaptureSonarScreenShot(url, filePath);
                //  SendEMailThroughOUTLOOK(filePath);

                if (ConfigurationManager.AppSettings["mailProtocol"].ToLower() != "smtp")
                {
                    SendEmailUsingOUTLOOK(filePath);
                }
                else
                    SendEmailUsingSMTP(filePath);
            }
            catch (Exception ex)
            {
                Log(ex.StackTrace);
                Log(ex.Message);
                //Log("Press any key to exit..");
                //Console.ReadLine();
            }
        }

        private static void CheckSonarAvailability()
        {
            try
            {
                string sonarUrl = ConfigurationManager.AppSettings["sonarUrl"];
                WebRequest request = WebRequest.Create(new Uri(sonarUrl));
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                if (response == null || response.StatusCode != HttpStatusCode.OK)
                {
                    Log("Sonarqube is not running.. Please start sonarqube and try again..");
                }
            }
            catch (Exception ex)
            {
                Log("Sonarqube is not running.. Please start sonarqube and try again..");
                throw ex;
            }
        }

        private static void CaptureSonarScreenShot(string url, string filePath)
        {
            try
            {
                Log("Capturing Sonar dashboard....\n");
                string render_string = "page.render('fileName');";
                string render_code = "";
                // string filePath = ConfigurationManager.AppSettings["screenCaptureSaveLocation"];
                //  filePath = filePath + DateTime.Now.ToString("yyyyMMdd") + "." + "pdf";

                render_code += Environment.NewLine + render_string.Replace("fileName", Path.Combine(filePath).Replace("\\", "/"));

                StreamReader reader = new StreamReader("Resources\\render_template.js");
                string source_content = reader.ReadToEnd();
                source_content = source_content.Replace("[UserURL]", url).Replace("[RENDER_CODE]", render_code);
                reader.Close();

                StreamWriter writer = new StreamWriter("render.js");
                writer.Write(source_content);
                writer.Close();

                var process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo { FileName = "phantomjs.exe", Arguments = "render.js" };
                process.Start();
                Thread.Sleep(60000);
                process.WaitForExit();
                
                Log($"Sonar Screen capture saved at {filePath}");
            }
            catch
            {
                Log($"Unable to Captrue and Save Sonar dashboard.. Please try again.");
                throw;
            }
        }

        public static string EncryptString(string plainText)
        {
               string initVector = "sonar99mckesson1234";
            string passPhrase = "121";
            int keysize = 256;
        byte[] initVectorBytes = Encoding.UTF8.GetBytes(initVector);
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);
            PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null);
            byte[] keyBytes = password.GetBytes(keysize / 8);
            RijndaelManaged symmetricKey = new RijndaelManaged();
            symmetricKey.Mode = CipherMode.CBC;
            ICryptoTransform encryptor = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes);
            MemoryStream memoryStream = new MemoryStream();
            CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write);
            cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
            cryptoStream.FlushFinalBlock();
            byte[] cipherTextBytes = memoryStream.ToArray();
            memoryStream.Close();
            cryptoStream.Close();
            return Convert.ToBase64String(cipherTextBytes);
        }
        //Decrypt
        public static string DecryptString(string cipherText)
        {
            string initVector = "sonar99mckesson1234";
            int keysize = 256;
            string passPhrase = "121";
            byte[] initVectorBytes = Encoding.UTF8.GetBytes(initVector);
            byte[] cipherTextBytes = Convert.FromBase64String(cipherText);
            PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null);
            byte[] keyBytes = password.GetBytes(keysize / 8);
            RijndaelManaged symmetricKey = new RijndaelManaged();
            symmetricKey.Mode = CipherMode.CBC;
            ICryptoTransform decryptor = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes);
            MemoryStream memoryStream = new MemoryStream(cipherTextBytes);
            CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];
            int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
            memoryStream.Close();
            cryptoStream.Close();
            return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount);
        }

    }
}