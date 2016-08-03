using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Data;
using System.Xml;
using System.IO;

namespace emailHelper
{
    public class emailHelper
    {
        public emailHelper() { }
        private string emailid;
        public string EmailID
        {
            get { return emailid; }
            set { emailid = value; }
        }
        private string emailpwd;
        public string EmailPWD
        {
            get { return emailpwd; }
            set { emailpwd = value; }
        }
        private string emailaddress;
        public string EmailAddress
        {
            get { return emailaddress; }
            set { emailaddress = value; }
        }
        private string emailsmtp;
        public string EmailSMTP
        {
            get { return emailsmtp; }
            set { emailsmtp = value; }
        }
        /// <summary>
        /// 以163邮箱发送邮件
        /// </summary>
        /// <param name="Subject">标题</param>
        /// <param name="Body">正文</param>
        /// <param name="TargetAddress">目标地址</param>
        /// <returns>发送成功</returns>
        public static string SendEmail(string Subject, string Body, string TargetAddress)
        {
            string strLocalAdd = ".\\Config.xml";
            emailHelper eh = new emailHelper();
            if (File.Exists(strLocalAdd))
            {
                try
                {
                    XmlDocument xmlCon = new XmlDocument();
                    xmlCon.Load(strLocalAdd);
                    XmlNode xnCon = xmlCon.SelectSingleNode("Config");
                    //string LinkString1 = xnCon.SelectSingleNode("LinkString1").InnerText;
                    //string LinkString2 = xnCon.SelectSingleNode("LinkString2").InnerText;                    
                    eh.EmailID= xnCon.SelectSingleNode("EmailID").InnerText;
                    eh.EmailPWD = xnCon.SelectSingleNode("EmailPWD").InnerText;
                    eh.EmailSMTP = xnCon.SelectSingleNode("EmailSMTP").InnerText;
                    eh.EmailAddress = xnCon.SelectSingleNode("EmailAddress").InnerText;
                }
                catch(Exception ex)
                {
                    return "Error:" + ex.ToString();
                }
            }
            else
            {
                return "Error:Config File Lost!";
            }

            string id = eh.EmailID;
            string pwd = eh.EmailPWD;
            string address = eh.EmailAddress;
            string smtp = eh.EmailSMTP;

            System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
            client.Host = smtp;//使用163的SMTP服务器发送邮件
            client.UseDefaultCredentials = true;
            client.DeliveryMethod = System.Net.Mail.SmtpDeliveryMethod.Network;
            client.Credentials = new System.Net.NetworkCredential(id, pwd);//163的SMTP服务器需要用163邮箱的用户名和密码作认证，如果没有需要去163申请个,
            System.Net.Mail.MailMessage Message = new System.Net.Mail.MailMessage();
            Message.From = new System.Net.Mail.MailAddress(address);//这里需要注意，163似乎有规定发信人的邮箱地址必须是163的，而且发信人的邮箱用户名必须和上面SMTP服务器认证时的用户名相同
            //因为上面用的用户名abc作SMTP服务器认证，所以这里发信人的邮箱地址也应该写为abc@163.com
            //Message.To.Add("123456@gmail.com");//将邮件发送给Gmail
            //Message.To.Add("12345@qq.com");//将邮件发送给QQ邮箱
            if (TargetAddress != "")
            {
                try
                {
                    Message.To.Add(TargetAddress);
                    Message.Subject = Subject;
                    Message.Body = Body;
                    Message.SubjectEncoding = System.Text.Encoding.UTF8;
                    Message.BodyEncoding = System.Text.Encoding.UTF8;
                    Message.Priority = System.Net.Mail.MailPriority.High;
                    Message.IsBodyHtml = true;
                    client.Send(Message);
                }
                catch (System.Net.Mail.SmtpException ex)
                {
                    return "Error:" + ex.ToString();
                }
            }
            return "Success!";
        }
    }
}
