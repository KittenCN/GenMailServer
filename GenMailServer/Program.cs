using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;
using System.Data;
using System.Threading;

namespace GenMailServer
{
    class Program
    {
        public static string GenLinkString = "./DB/GenMailServer.accdb";
        public static string GenCheckStr = "MailQueues";
        public static string LinkCheckStr = "MailTrans";
        public static string LinkString1;
        public static string LinkString2;
        public static int EmailRete = 30;
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Timer t = new Timer(TimerCallback, null, 0, 60000);

            bool boolstatus = false;
            Console.WriteLine("Welcome to GMS-General Mail Server");
            Console.WriteLine("Reading Config File ...");
            string strLocalAdd = ".\\Config.xml";
            emailHelper.emailHelper eh = new emailHelper.emailHelper();
            if (File.Exists(strLocalAdd))
            {
                try
                {
                    XmlDocument xmlCon = new XmlDocument();
                    xmlCon.Load(strLocalAdd);
                    XmlNode xnCon = xmlCon.SelectSingleNode("Config");
                    GenLinkString = "./DB/GenMailServer.accdb";
                    GenCheckStr = "MailQueues";
                    LinkCheckStr = "MailTrans";
                    LinkString1 = xnCon.SelectSingleNode("LinkString1").InnerText;
                    LinkString2 = xnCon.SelectSingleNode("LinkString2").InnerText;
                    EmailRete = int.Parse(xnCon.SelectSingleNode("EmailRate").InnerText);
                    if(AccessHelper.AccessHelper.CheckDB(GenLinkString,GenCheckStr) && AccessHelper.AccessHelper.CheckDB(LinkString1,LinkCheckStr) && AccessHelper.AccessHelper.CheckDB(LinkString2, LinkCheckStr))
                    {
                        boolstatus = true;
                    }
                    else
                    {
                        boolstatus = false;
                    }
                }
                catch (Exception ex)
                {
                    boolstatus = false;
                    Console.WriteLine("Error:" + ex.ToString());
                }
            }
            else
            {
                boolstatus = false;
                Console.WriteLine("Error:Config File Lost!");
            }
            if(boolstatus)
            {
                try
                {
                    Console.WriteLine("Trans Data to Local DB from LinkString1...");
                    AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(LinkString1);
                    string strSQL = "select * from " + LinkCheckStr;
                    DataTable dtSQL = ah.ReturnDataTable(strSQL);
                    TransToLocal(dtSQL);
                    strSQL = "delete from " + LinkCheckStr;
                    ah.ExecuteNonQuery(strSQL);

                    Console.WriteLine("Trans Data to Local DB from LinkString1...");
                    ah = new AccessHelper.AccessHelper(LinkString2);
                    strSQL = "select * from " + LinkCheckStr;
                    dtSQL = ah.ReturnDataTable(strSQL);
                    TransToLocal(dtSQL);
                    strSQL = "delete from " + LinkCheckStr;
                    ah.ExecuteNonQuery(strSQL);

                    Console.WriteLine("Processing the Local Mail Queues...");
                    ah = new AccessHelper.AccessHelper(GenLinkString);
                    strSQL = "select * from " + GenCheckStr;
                    dtSQL = ah.ReturnDataTable(strSQL);
                    int i = 0;
                    foreach(DataRow row in dtSQL.Rows)
                    {
                        i++;
                        string strMailResult = emailHelper.emailHelper.SendEmail(row["MailSubject"].ToString(), row["MailBody"].ToString(), row["MailTargetAddress"].ToString());
                        if(strMailResult== "Success!")
                        {
                            Console.WriteLine("The " + i + " Mail has been sent successfully!");
                            string strInSQL = "insert into MailHistory(MailSubject,MailBody,MailTargetAddress,MailDateTime,SendDateTime) ";
                            strInSQL = strInSQL + " values('" + row["MailSubject"].ToString() + "','" + row["MailBody"].ToString() + "','" + row["MailTargetAddress"].ToString() + "',#" + row["MailDateTime"].ToString() + "#,#" + DateTime.Now.ToString() + "#) ";
                            ah.ExecuteNonQuery(strInSQL);
                            strInSQL = "delete from " + GenCheckStr + " where id=" + row["ID"].ToString() + " ";
                            ah.ExecuteNonQuery(strInSQL);
                            Console.WriteLine("The " + i + " Mail has been processed successfully!");
                            Thread.Sleep(EmailRete * 1000);
                        }
                        else
                        {
                            Console.WriteLine("Error:" + strMailResult);
                        }
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error:" + ex.ToString());
                }
            }
            else
            {
                Console.WriteLine("Error:" + "Some Boolean Values is False!");
            }
            Console.ReadLine();
        }
        private static void TransToLocal(DataTable dt)
        {
            AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(GenLinkString);
            foreach(DataRow row in dt.Rows)
            {
                string strSQL = "insert into MailQueues(MailSubject,MailBody,MailTargetAddress,MailDateTime) ";
                strSQL = strSQL + " values('" + row["MailSubject"].ToString() + "','" + row["MailBody"].ToString() + "','" + row["MailTargetAddress"].ToString() + "',#" + DateTime.Now.ToString() + "#) ";
                ah.ExecuteNonQuery(strSQL); 
            }
        }

        private static void TimerCallback(Object o)
        {
            // Display the date/time when this method got called.
            Console.WriteLine("In TimerCallback: " + DateTime.Now);
            // Force a garbage collection to occur for this demo.
            GC.Collect();
        }
    }
}
