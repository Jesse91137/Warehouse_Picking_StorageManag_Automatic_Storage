using System;
using System.Net;
using System.Net.Mail;

namespace Automatic_Storage
{
    /// <summary>
    /// 提供 SMTP 郵件發送功能的類別。
    /// </summary>
    public class sendMail
    {
        /// <summary>
        /// SMTP 伺服器位址。
        /// </summary>
        private string smtpServer;
        /// <summary>
        /// SMTP 伺服器連接埠。
        /// </summary>
        private int smtpPort;
        /// <summary>
        /// SMTP 登入帳號。
        /// </summary>
        private string smtpUser;
        /// <summary>
        /// SMTP 登入密碼。
        /// </summary>
        private string smtpPass;

        /// <summary>
        /// 建構函式，初始化 SMTP 設定。
        /// </summary>
        /// <param name="server">SMTP 伺服器位址</param>
        /// <param name="port">SMTP 連接埠</param>
        /// <param name="user">SMTP 使用者名稱</param>
        /// <param name="pass">SMTP 密碼</param>
        public sendMail(string server, int port, string user, string pass)
        {
            smtpServer = server; // 設定 SMTP 伺服器位址
            smtpPort = port;     // 設定 SMTP 連接埠
            smtpUser = user;     // 設定 SMTP 使用者名稱
            smtpPass = pass;     // 設定 SMTP 密碼
        }

        /// <summary>
        /// 發送電子郵件。
        /// </summary>
        /// <param name="fromEmail">寄件者電子郵件</param>
        /// <param name="toEmails">收件者電子郵件（多個以分號分隔）</param>
        /// <param name="subject">郵件主旨</param>
        /// <param name="body">郵件內容</param>
        /// <exception cref="ApplicationException">發送失敗時拋出例外</exception>
        public void SendEmail(string fromEmail, string toEmails, string subject, string body)
        {
            try
            {
                MailMessage mail = new MailMessage(); // 建立郵件訊息物件
                SmtpClient smtpClient = new SmtpClient(smtpServer); // 建立 SMTP 用戶端

                mail.From = new MailAddress(fromEmail); // 設定寄件者
                string[] recipients = toEmails.Split(';'); // 以分號分割收件者
                foreach (string recipient in recipients)
                {
                    if (!string.IsNullOrEmpty(recipient)) // 檢查收件者是否為空
                    {
                        mail.To.Add(new MailAddress(recipient)); // 加入收件者
                    }
                }

                mail.Subject = subject; // 設定郵件主旨
                mail.Body = body;       // 設定郵件內容
                mail.IsBodyHtml = true; // 設定郵件內容為 HTML 格式

                smtpClient.Port = smtpPort; // 設定 SMTP 連接埠
                smtpClient.Credentials = new NetworkCredential(smtpUser, smtpPass); // 設定 SMTP 認證
                smtpClient.EnableSsl = true; // 啟用 SSL 加密

                smtpClient.Send(mail); // 發送郵件
            }
            catch (SmtpException ex)
            {
                // 處理 SMTP 相關錯誤
                throw new ApplicationException("SMTP error occurred: " + ex.Message);
            }
            catch (Exception ex)
            {
                // 處理其他錯誤
                throw new ApplicationException("An error occurred: " + ex.Message);
            }
        }
    }
}
