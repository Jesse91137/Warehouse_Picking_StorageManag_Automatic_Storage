using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Automatic_Storage
{
    public class sendMail
    {
        private string smtpServer;
        private int smtpPort;
        private string smtpUser;
        private string smtpPass;

        // Constructor to initialize the SMTP settings
        public sendMail(string server, int port, string user, string pass)
        {
            smtpServer = server;
            smtpPort = port;
            smtpUser = user;
            smtpPass = pass;
        }

        // Method to send an email
        public void SendEmail(string fromEmail, string toEmails, string subject, string body)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient smtpClient = new SmtpClient(smtpServer);

                mail.From = new MailAddress(fromEmail);
                string[] recipients = toEmails.Split(';');
                foreach (string recipient in recipients)
                {
                    if (!string.IsNullOrEmpty(recipient))
                    {
                        mail.To.Add(new MailAddress(recipient));
                    }
                }

                mail.Subject = subject;
                mail.Body = body;
                mail.IsBodyHtml = true; // Set to false if the email is plain text

                smtpClient.Port = smtpPort;
                smtpClient.Credentials = new NetworkCredential(smtpUser, smtpPass);
                smtpClient.EnableSsl = true;

                smtpClient.Send(mail);
            }
            catch (SmtpException ex)
            {
                // Handle the error (e.g., log it or show a message to the user)
                throw new ApplicationException("SMTP error occurred: " + ex.Message);
            }
            catch (Exception ex)
            {
                // Handle other errors
                throw new ApplicationException("An error occurred: " + ex.Message);
            }
        }
    }
}
