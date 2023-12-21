using System.Net.Mail;
using System.Net;
using System.Text;
using System.Net.Mime;

namespace Working.Tools
{
    /// <summary>
    /// 发送邮件类
    /// </summary>
    public class EmailSender
    {
        private string smtpServer;
        private int smtpPort;
        private string senderName;
        private string senderPassword;
        private string recipientEmail;
        private string subjectPrefix;
        private string emailContext;

        /// <summary>
        /// 发送邮件类
        /// </summary>
        /// <param name="smtpServer">smtp服务器地址</param>
        /// <param name="smtpPort">smtp端口</param>
        /// <param name="senderName">发件人邮箱地址</param>
        /// <param name="senderPassword">发件人密码</param>
        /// <param name="recipientEmail">收件人邮箱地址</param>
        /// <param name="subjectPrefix">主题</param>
        /// <param name="emailContext">内容</param>
        public EmailSender(string smtpServer, int smtpPort, string senderName, string senderPassword, string recipientEmail, string subjectPrefix, string emailContext)
        {
            this.smtpServer = smtpServer;
            this.smtpPort = smtpPort;
            this.senderName = senderName;
            this.senderPassword = senderPassword;
            this.recipientEmail = recipientEmail;
            this.subjectPrefix = subjectPrefix;
            this.emailContext = emailContext;
        }
        ///
        public void SendMail(string attachmentFilePath="")
        {
            try
            {
                SmtpClient smtp = new SmtpClient(smtpServer, smtpPort);
                smtp.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
                smtp.Credentials = new NetworkCredential(senderName, senderPassword);

                MailMessage message = new MailMessage();
                message.From = new MailAddress(senderName);
                message.To.Add(new MailAddress(recipientEmail));
                message.Subject = $"{subjectPrefix} ";
                message.Body = $"{emailContext}";
                message.BodyEncoding = Encoding.UTF8;
                message.IsBodyHtml = true;

                if (File.Exists(attachmentFilePath))
                {
                    // 添加附件  
                    Attachment attachment = new Attachment(attachmentFilePath, MediaTypeNames.Application.Octet);
                    attachment.ContentDisposition.CreationDate = File.GetCreationTime(attachmentFilePath);
                    attachment.ContentDisposition.ModificationDate = File.GetLastWriteTime(attachmentFilePath);
                    attachment.ContentDisposition.ReadDate = File.GetCreationTime(attachmentFilePath);
                    message.Attachments.Add(attachment);
                }          
                smtp.Send(message);
                Console.WriteLine("发送成功！");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }



        /*调用示例*/
        /*
            string smtpServer = "smtp.qq.com"; // 或者 "smtp.163.com"  
            int smtpPort = 587; // 根据SMTP服务器配置更改这个值  网易邮箱端口 25
            string senderName = "1111@qq.com"; // 发件人邮箱地址  
            string senderPassword = "mzodpkjdipvabgibc"; // 发件人邮箱密码  
            string recipientEmail = "17600813451@163.com"; // 收件人邮箱地址  
            string subjectPrefix = $"{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}  测试邮件标题";// 邮件主题前缀，可以根据需要更改这个值  
            string attachmentFilePath = ""; // 附件文件路径，根据需要更改这个值  
            string emailContext = $"{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}  测试邮件内容";

            EmailSender emailSender = new EmailSender(smtpServer, smtpPort, senderName, senderPassword, recipientEmail, subjectPrefix, emailContext);
            emailSender.SendMail(attachmentFilePath); // 替换为你需要发送的邮件内容和文件路径，包括附件和图片的路径
         
         
         */
    }
}
