using System;
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
        public static int EmailRete = 10;
        public static string strLocalAdd = ".\\Config.xml";
        public static Boolean boolClockShow = false;
        public static Timer t;
        public static Timer tClock;
        public static Boolean boolProcess = false;
        public static int intMainRate = 60;
        public static int intSecondShow = 60;
        public static int intEmailTestFlag = 0;
        public static string strEmailTestAddress = "owdely@163.com";
        public static Boolean boolSilentTimeShow = false;
        public static int intSilentTime = 10;        
        static void Main(string[] args)
        {
            ConsoleHelper.ConsoleHelper.cInitiaze();
            ConsoleHelper.ConsoleHelper.wl("Welcome to GMS-General Mail Server");
            ConsoleHelper.ConsoleHelper.wl("");
            for(int i=0;i<10;i++)
            {
                if(i%2==0)
                {                   
                    ConsoleHelper.ConsoleHelper.wrr("DO NOT CLICK THE INTERFACE AND DO NOT PRESS ANYKEY WHEN THE RUNNING CONSOLE IS IN THE FOREGROUND!!",ConsoleColor.Red, ConsoleColor.Cyan);
                    Thread.Sleep(500);
                }
                else
                {
                    ConsoleHelper.ConsoleHelper.wrr("DO NOT CLICK THE INTERFACE AND DO NOT PRESS ANYKEY WHEN THE RUNNING CONSOLE IS IN THE FOREGROUND!!",ConsoleColor.DarkRed, ConsoleColor.Cyan);
                    Thread.Sleep(500);
                }
            }
            ConsoleHelper.ConsoleHelper.wrr("DO NOT CLICK THE INTERFACE AND DO NOT PRESS ANYKEY WHEN THE RUNNING CONSOLE IS IN THE FOREGROUND!!", ConsoleColor.Red, ConsoleColor.Cyan);
            ConsoleHelper.ConsoleHelper.wl("");
            ConsoleHelper.ConsoleHelper.wl("");
            ConsoleHelper.ConsoleHelper.cInitiaze();
            if (File.Exists(strLocalAdd))
            {
                try
                {
                    ConsoleHelper.ConsoleHelper.wl("Reading Config File ...");
                    XmlDocument xmlCon = new XmlDocument();
                    xmlCon.Load(strLocalAdd);
                    XmlNode xnCon = xmlCon.SelectSingleNode("Config");
                    GenLinkString = "./DB/GenMailServer.accdb";
                    GenCheckStr = "MailQueues";
                    LinkCheckStr = "MailTrans";
                    LinkString1 = xnCon.SelectSingleNode("LinkString1").InnerText;
                    LinkString2 = xnCon.SelectSingleNode("LinkString2").InnerText;
                    EmailRete = int.Parse(xnCon.SelectSingleNode("EmailRate").InnerText);
                    intEmailTestFlag = int.Parse(xnCon.SelectSingleNode("EmailTestFlag").InnerText);
                    intMainRate=int.Parse(xnCon.SelectSingleNode("MainRate").InnerText);
                    strEmailTestAddress = xnCon.SelectSingleNode("EmailTestAddress").InnerText;
                    ConsoleHelper.ConsoleHelper.wl("Reading Config File Successfully...");

                    intSilentTime = EmailRete;
                    intSecondShow = intMainRate;
                    ConsoleHelper.ConsoleHelper.wl("Begin Timer Methods...");
                    t = new Timer(TimerCallback, null, 0, intMainRate * 1000);
                    tClock = new Timer(TimerClockShow, null, 0, 1000);
                }
                catch (Exception ex)
                {
                    ConsoleHelper.ConsoleHelper.wl("Error:" + ex.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                }
            }
            else
            {
                ConsoleHelper.ConsoleHelper.wl("Error:Config File Lost!", ConsoleColor.Red, ConsoleColor.Black);
            }
            Console.ReadLine();
        }
        private static void TransToLocal(DataTable dt,int intFlag)
        {
            AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(GenLinkString);
            foreach (DataRow row in dt.Rows)
            {
                string strSQL = "insert into MailQueues(MailSubject,MailBody,MailTargetAddress,MailDateTime,Flag) ";
                strSQL = strSQL + " values('" + row["MailSubject"].ToString() + "','" + row["MailBody"].ToString() + "','" + row["MailTargetAddress"].ToString() + "',#" + DateTime.Now.ToString() + "#," + intFlag + ") ";
                ah.ExecuteNonQuery(strSQL);
            }
        }
         
        private static void TimerClockShow(object o)
        {
            if (intSecondShow > 0)
            {
                intSecondShow--;
            }
            else
            {
                intSecondShow = intMainRate;
            }
            //if (intSilentTime > 0)
            //{
            //    intSilentTime--;
            //}
            //else
            //{
            //    intSilentTime = EmailRete;
            //}
            if (!boolProcess)
            {
                if (!boolClockShow)
                {
                    ConsoleHelper.ConsoleHelper.wl("");
                    ConsoleHelper.ConsoleHelper.wrr("Now is :" + DateTime.Now.ToString() + " ...");
                    boolClockShow = true;
                }
                else
                {
                    ConsoleHelper.ConsoleHelper.wrr("Now is :" + DateTime.Now.ToString() + " , and " + intSecondShow + " seconds to the next execution.");
                }
            }
            //else if(boolProcess && boolSilentTimeShow)
            //{
            //    Console.WriteLine("");
            //    Console.Write("\rSilent Time : " + intSilentTime + " Sec Left...");
            //}
            GC.Collect();
        }

        private static void TimerCallback(Object o)
        {
            // Display the date/time when this method got called.
            //Console.WriteLine("In TimerCallback: " + DateTime.Now);
            // Force a garbage collection to occur for this demo.
            ConsoleHelper.ConsoleHelper.wl("");
            ConsoleHelper.ConsoleHelper.wl("");
            ConsoleHelper.ConsoleHelper.wl("Running Main Method...");

            //boolClockShow = false;
            boolProcess = true;
            bool boolstatus = false;
            emailHelper.emailHelper eh = new emailHelper.emailHelper();
            try
            {
                if (intEmailTestFlag == 1)
                {
                    ConsoleHelper.ConsoleHelper.wl("Debug Mode Open...");
                    for(int i=1; i<=2; i++)
                    {
                        string strDebugResult = emailHelper.emailHelper.SendEmail("TestSubject", "TestBody", strEmailTestAddress, i);
                        if (strDebugResult == "Success!")
                        {
                            ConsoleHelper.ConsoleHelper.wl("The Debug Mail With LinkString[" + i + "] has been sent successfully!");
                            Thread.Sleep(EmailRete * 1000);
                        }
                        else
                        {
                            ConsoleHelper.ConsoleHelper.wl("LinkString[" + i + "] had Error:" + strDebugResult, ConsoleColor.Red, ConsoleColor.Black);
                        }
                    }
                }
                if (AccessHelper.AccessHelper.CheckDB(GenLinkString, GenCheckStr) && AccessHelper.AccessHelper.CheckDB(LinkString1, LinkCheckStr) && AccessHelper.AccessHelper.CheckDB(LinkString2, LinkCheckStr))
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
                ConsoleHelper.ConsoleHelper.wl("Error:" + ex.ToString(), ConsoleColor.Red, ConsoleColor.Black);
            }
            if (boolstatus)
            {
                try
                {
                    ConsoleHelper.ConsoleHelper.wl("Trans Data to Local DB from LinkString1...");
                    AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(LinkString1);
                    string strSQL = "select * from " + LinkCheckStr;
                    DataTable dtSQL = ah.ReturnDataTable(strSQL);
                    TransToLocal(dtSQL,1);
                    strSQL = "delete from " + LinkCheckStr;
                    ah.ExecuteNonQuery(strSQL);

                    ConsoleHelper.ConsoleHelper.wl("Trans Data to Local DB from LinkString2...");
                    ah = new AccessHelper.AccessHelper(LinkString2);
                    strSQL = "select * from " + LinkCheckStr;
                    dtSQL = ah.ReturnDataTable(strSQL);
                    TransToLocal(dtSQL,2);
                    strSQL = "delete from " + LinkCheckStr;
                    ah.ExecuteNonQuery(strSQL);

                    ConsoleHelper.ConsoleHelper.wl("Processing the Local Mail Queues...");
                    ah = new AccessHelper.AccessHelper(GenLinkString);
                    strSQL = "select * from " + GenCheckStr;
                    dtSQL = ah.ReturnDataTable(strSQL);
                    int i = 0;
                    foreach (DataRow row in dtSQL.Rows)
                    {
                        i++;
                        string strMailResult = emailHelper.emailHelper.SendEmail(row["MailSubject"].ToString(), row["MailBody"].ToString(), row["MailTargetAddress"].ToString(),int.Parse(row["Flag"].ToString()));
                        if (strMailResult == "Success!")
                        {
                            ConsoleHelper.ConsoleHelper.wl("The " + i + " Mail has been sent successfully!");
                            string strInSQL = "insert into MailHistory(MailSubject,MailBody,MailTargetAddress,MailDateTime,SendDateTime,Flag) ";
                            strInSQL = strInSQL + " values('" + row["MailSubject"].ToString() + "','" + row["MailBody"].ToString() + "','" + row["MailTargetAddress"].ToString() + "',#" + row["MailDateTime"].ToString() + "#,#" + DateTime.Now.ToString() + "#," + int.Parse(row["Flag"].ToString()) + ") ";
                            ah.ExecuteNonQuery(strInSQL);
                            strInSQL = "delete from " + GenCheckStr + " where id=" + row["ID"].ToString() + " ";
                            ah.ExecuteNonQuery(strInSQL);
                            ConsoleHelper.ConsoleHelper.wl("The " + i + " Mail has been processed successfully!");
                            Thread.Sleep(EmailRete * 1000);
                        }
                        else
                        {
                            ConsoleHelper.ConsoleHelper.wl("Error:" + strMailResult, ConsoleColor.Red, ConsoleColor.Black);
                        }
                    }
                }
                catch (Exception ex)
                {
                    ConsoleHelper.ConsoleHelper.wl("Error:" + ex.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                }
            }
            else
            {
                ConsoleHelper.ConsoleHelper.wl("Error:" + "Some Boolean Values is False!",ConsoleColor.Red,ConsoleColor.Black);
            }
            boolProcess = false;
            ConsoleHelper.ConsoleHelper.wl("End Running...");
            ConsoleHelper.ConsoleHelper.wl("");
            GC.Collect();
        }
    }
}
