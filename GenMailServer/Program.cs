﻿using System;
using System.IO;
using System.Xml;
using System.Data;
using System.Threading;

namespace GenMailServer
{
    class Program
    {
        #region 全局变量
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
        public static Timer tDBCache;
        public static Boolean boolProcess = false;
        public static int intMainRate = 60;
        public static int intSecondShow = 60;
        public static int intEmailTestFlag = 0;
        public static string strEmailTestAddress = "owdely@163.com";
        public static Boolean boolSilentTimeShow = false;
        public static int intSilentTime = 10;
        public static Boolean boolDBCache = false;
        public static int intDBCacheRate = 3600;
        public static int int3rdShow = 60;
        #endregion
        #region Main Method
        static void Main(string[] args)
        {
            CheckDB(GenLinkString, GenCheckStr);
            CheckDB(LinkString1, LinkCheckStr);
            CheckDB(LinkString2, LinkCheckStr);

            ConsoleHelper.ConsoleHelper.cInitiaze();
            ConsoleHelper.ConsoleHelper.wl("Welcome to GMS-General Mail Server");
            ConsoleHelper.ConsoleHelper.wl("");
            for (int i = 0; i < 10; i++)
            {
                if (i % 2 == 0)
                {
                    ConsoleHelper.ConsoleHelper.wrr("DO NOT CLICK THE INTERFACE AND DO NOT PRESS ANYKEY WHEN THE RUNNING CONSOLE IS IN THE FOREGROUND!!", ConsoleColor.Red, ConsoleColor.Cyan, false);
                    Thread.Sleep(500);
                }
                else
                {
                    ConsoleHelper.ConsoleHelper.wrr("DO NOT CLICK THE INTERFACE AND DO NOT PRESS ANYKEY WHEN THE RUNNING CONSOLE IS IN THE FOREGROUND!!", ConsoleColor.DarkRed, ConsoleColor.Cyan, false);
                    Thread.Sleep(500);
                }
            }
            ConsoleHelper.ConsoleHelper.wrr("DO NOT CLICK THE INTERFACE AND DO NOT PRESS ANYKEY WHEN THE RUNNING CONSOLE IS IN THE FOREGROUND!!", ConsoleColor.Red, ConsoleColor.Cyan, true);
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
                    intMainRate = int.Parse(xnCon.SelectSingleNode("MainRate").InnerText);
                    strEmailTestAddress = xnCon.SelectSingleNode("EmailTestAddress").InnerText;
                    intDBCacheRate = int.Parse(xnCon.SelectSingleNode("DBCacheRate").InnerText);
                    ConsoleHelper.ConsoleHelper.wl("Reading Config File Successfully...");

                    intSilentTime = EmailRete;
                    intSecondShow = intMainRate;
                    int3rdShow = intDBCacheRate;
                    TimerCallback(null);
                    TimerDBCacheProcess(null);
                    ConsoleHelper.ConsoleHelper.wl("Begin Timer Methods...");
                    tClock = new Timer(TimerClockShow, null, 0, 1000);
                    t = new Timer(TimerCallback, null, 0, intMainRate * 1000);                    
                    tDBCache = new Timer(TimerDBCacheProcess, null, 0, intDBCacheRate * 1000);
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
            ConsoleHelper.ConsoleHelper.wl("Exit!");
        }
        #endregion
        #region 提取远端邮件数据到本地
        private static void TransToLocal(DataTable dt, int intFlag)
        {
            AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(GenLinkString);
            foreach (DataRow row in dt.Rows)
            {
                string strSQL = "insert into MailQueues(MailSubject,MailBody,MailTargetAddress,MailDateTime,Flag) ";
                strSQL = strSQL + " values('" + row["MailSubject"].ToString() + "','" + row["MailBody"].ToString() + "','" + row["MailTargetAddress"].ToString() + "',#" + DateTime.Now.ToString() + "#," + intFlag + ") ";
                ah.ExecuteNonQuery(strSQL);
            }
        }
        #endregion
        #region 时间显示事件
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
            if (int3rdShow > 0)
            {
                int3rdShow--;
            }
            else
            {
                int3rdShow = intDBCacheRate;
            }
            //if (intSilentTime > 0)
            //{
            //    intSilentTime--;
            //}
            //else
            //{
            //    intSilentTime = EmailRete;
            //}
            if (!boolProcess || !boolDBCache)
            {
                if (!boolClockShow)
                {
                    ConsoleHelper.ConsoleHelper.wl("");
                    ConsoleHelper.ConsoleHelper.wrr("Now is :" + DateTime.Now.ToString() + " ...", true);
                    boolClockShow = true;
                }
                else
                {
                    int intShow = 999;
                    if (intSecondShow > int3rdShow)
                    {
                        intShow = int3rdShow;
                    }
                    else
                    {
                        intShow = intSecondShow;
                    }
                    ConsoleHelper.ConsoleHelper.wrr("Now is :" + DateTime.Now.ToString() + " , and " + intShow + " seconds to the next execution.", false);
                }
            }
            //else if(boolProcess && boolSilentTimeShow)
            //{
            //    Console.WriteLine("");
            //    Console.Write("\rSilent Time : " + intSilentTime + " Sec Left...");
            //}
            GC.Collect();
        }
        #endregion
        #region 邮件处理事件
        private static void TimerCallback(Object o)
        {
            if (!boolDBCache && !boolProcess)
            {
                // Display the date/time when this method got called.
                //Console.WriteLine("In TimerCallback: " + DateTime.Now);
                // Force a garbage collection to occur for this demo.
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
                        for (int i = 1; i <= 2; i++)
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
                    if (CheckDB(GenLinkString, GenCheckStr) && CheckDB(LinkString1, LinkCheckStr) && CheckDB(LinkString2, LinkCheckStr))
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
                        TransToLocal(dtSQL, 1);
                        strSQL = "delete from " + LinkCheckStr;
                        ah.ExecuteNonQuery(strSQL);

                        ConsoleHelper.ConsoleHelper.wl("Trans Data to Local DB from LinkString2...");
                        ah = new AccessHelper.AccessHelper(LinkString2);
                        strSQL = "select * from " + LinkCheckStr;
                        dtSQL = ah.ReturnDataTable(strSQL);
                        TransToLocal(dtSQL, 2);
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
                            string strMailResult = emailHelper.emailHelper.SendEmail(row["MailSubject"].ToString(), row["MailBody"].ToString(), row["MailTargetAddress"].ToString(), int.Parse(row["Flag"].ToString()));
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
                    ConsoleHelper.ConsoleHelper.wl("Error:" + "Local DB or Remote DB can not be connected!", ConsoleColor.Red, ConsoleColor.Black);
                }
                boolProcess = false;
                ConsoleHelper.ConsoleHelper.wl("End Running...");
                ConsoleHelper.ConsoleHelper.wl("");
            }
            GC.Collect();
        }
        #endregion
        #region 数据库及数据表测试
        public static Boolean CheckDB(string strDBAddress, string strTableName)
        {
            Boolean boolResult = false;
            AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(strDBAddress);
            if (ah.ConnectTest())
            {
                try
                {
                    string strSQL = "select * from " + strTableName;
                    DataTable dtSQL = ah.ReturnDataTable(strSQL);
                    if (dtSQL.Rows.Count >= 0)
                    {
                        boolResult = true;
                    }
                    else
                    {
                        boolResult = false;
                    }
                }
                catch (Exception ex1)
                {
                    if (ex1.HResult.ToString() == "-2147217865" && strTableName == "MailTrans")
                    {
                        try
                        {
                            ConsoleHelper.ConsoleHelper.wl("Can not found MailTrans table , system will try to create it...");
                            string strInSQL = "create table MailTrans(id autoincrement,MailSubject longtext,MailBody longtext,MailTargetAddress longtext,Flag int)";
                            ConsoleHelper.ConsoleHelper.wl("Create the MailTrans table successfully.");
                            ah.ExecuteNonQuery(strInSQL);
                            boolResult = true;
                        }
                        catch (Exception ex2)
                        {
                            ConsoleHelper.ConsoleHelper.wl("Error:" + ex2.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                            boolResult = false;
                        }
                    }
                    else
                    {
                        ConsoleHelper.ConsoleHelper.wl("Error:" + ex1.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                        boolResult = false;
                    }
                }
                if (strTableName == GenCheckStr)
                {
                    try
                    {
                        string strSQL = "select * from Log";
                        DataTable dtSQL = ah.ReturnDataTable(strSQL);
                        if (dtSQL.Rows.Count >= 0)
                        {
                            boolResult = true;
                        }
                        else
                        {
                            boolResult = false;
                        }
                    }
                    catch (Exception ex1)
                    {
                        if (ex1.HResult.ToString() == "-2147217865")
                        {
                            try
                            {
                                ConsoleHelper.ConsoleHelper.wl("Can not found Log table , system will try to create it...", false);
                                string strInSQL = "create table Log(id autoincrement,Log longtext,LogDateTime datetime)";
                                ah.ExecuteNonQuery(strInSQL);
                                boolResult = true;
                                ConsoleHelper.ConsoleHelper.wl("Create the Log table successfully.");
                            }
                            catch (Exception ex2)
                            {
                                ConsoleHelper.ConsoleHelper.wl("Error:" + ex2.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                                boolResult = false;
                            }
                        }
                        else
                        {
                            ConsoleHelper.ConsoleHelper.wl("Error:" + ex1.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                            boolResult = false;
                        }
                    }
                }
                else if (strTableName == LinkCheckStr)
                {
                    try
                    {
                        string strSQL = "select PhoneNUM from ApplicationDetail";
                        DataTable dtSQL = ah.ReturnDataTable(strSQL);
                        if (dtSQL.Rows.Count >= 0)
                        {
                            boolResult = true;
                        }
                        else
                        {
                            boolResult = false;
                        }
                    }
                    catch (Exception ex1)
                    {
                        if (ex1.HResult.ToString() == "-2147217904")
                        {
                            try
                            {
                                ConsoleHelper.ConsoleHelper.wl("Can not found PhonNUM String , system will try to create it...");
                                string strInSQL = "alter table ApplicationDetail add COLUMN PhoneNUM text";
                                ah.ExecuteNonQuery(strInSQL);
                                boolResult = true;
                                ConsoleHelper.ConsoleHelper.wl("Create the PhonNUM String successfully.");
                            }
                            catch (Exception ex2)
                            {
                                ConsoleHelper.ConsoleHelper.wl("Error:" + ex2.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                                boolResult = false;
                            }
                        }
                        else
                        {
                            ConsoleHelper.ConsoleHelper.wl("Error:" + ex1.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                            boolResult = false;
                        }
                    }
                }
            }
            else
            {
                boolResult = false;
            }
            return boolResult;
        }
        #endregion
        #region 远程数据缓存处理事件
        private static void TimerDBCacheProcess(object o)
        {
            if (!boolDBCache && !boolProcess)
            {
                boolDBCache = true;
                ConsoleHelper.ConsoleHelper.wl("");
                ConsoleHelper.ConsoleHelper.wl("Running DB Cache Method...");
                ConsoleHelper.ConsoleHelper.wl("Checking DB Cache Directory...");
                int intSuccess = 0;
                int intError = 0;
                string url = LinkString2.Substring(0, LinkString2.LastIndexOf("\\") + 1) + "DBCache\\";
                try
                {
                    if (Directory.Exists(url))
                    {
                        DirectoryInfo di = new DirectoryInfo(url);
                        ConsoleHelper.ConsoleHelper.wl("Checking DB Cache Files...");
                        foreach (FileInfo fi in di.GetFiles("*.accdb"))
                        {
                            AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(fi.FullName);
                            if (ah.ConnectTest())
                            {
                                ConsoleHelper.ConsoleHelper.wl("DB Cache Main Method...");
                                string strSQL = "select * from AccessQueue";
                                DataTable dtSQL = ah.ReturnDataTable(strSQL);
                                string strLastCtrlID = "";
                                Boolean boolLastCtrlID = true; ;
                                foreach (DataRow dr in dtSQL.Rows)
                                {
                                    if (dr["TransNo"] != null && dr["TransNo"].ToString() != "" && dr["operation"].ToString().Substring(0, 5) == "Cache")
                                    {
                                        string strUID = dr["operation"].ToString().Substring(5);
                                        if (strLastCtrlID == dr["TransNo"].ToString() && boolLastCtrlID == false)
                                        {
                                            intError++;
                                            break;
                                        }
                                        else
                                        {
                                            strLastCtrlID = dr["TransNo"].ToString();
                                            if (GetTotalPricefromUID(strUID) - double.Parse(dr["Buy"].ToString()) >= 0 || dr["Buy"].ToString() == "" || double.Parse(dr["Buy"].ToString()) == 0)
                                            {
                                                AccessHelper.AccessHelper ahLink2 = new AccessHelper.AccessHelper(LinkString2);
                                                if (dr["SqlStr"] != null && dr["SqlStr"].ToString() != "")
                                                {
                                                    ahLink2.ExecuteNonQuery(dr["SqlStr"].ToString());
                                                    string strSQLin = "delete from AccessQueue where ID=" + dr["ID"].ToString() + " ";
                                                    ah.ExecuteNonQuery(strSQLin);
                                                    intSuccess++;
                                                }
                                            }
                                            else
                                            {
                                                string strSQLin = "";
                                                AccessHelper.AccessHelper ahLink2 = new AccessHelper.AccessHelper(LinkString2);
                                                if (dr["SqlStr"] != null && dr["SqlStr"].ToString() != "")
                                                {
                                                    ahLink2.ExecuteNonQuery(dr["SqlStr"].ToString());
                                                    strSQLin = "update ApplicationInfo set AppState=-1 where TransNo='" + dr["TransNo"].ToString() + "' ";
                                                    ahLink2.ExecuteNonQuery(strSQLin);
                                                }
                                                strSQLin = "delete from AccessQueue where ID=" + dr["ID"].ToString() + " ";
                                                ah.ExecuteNonQuery(strSQLin);
                                                intError++;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        intError++;
                                    }
                                }
                            }
                        }
                        if (intSuccess > 0 || intError > 0)
                        {
                            ConsoleHelper.ConsoleHelper.wl("");
                            ConsoleHelper.ConsoleHelper.wl("Result:");
                            ConsoleHelper.ConsoleHelper.wl("Success:" + intSuccess);
                            ConsoleHelper.ConsoleHelper.wl("Error:" + intError);
                            ConsoleHelper.ConsoleHelper.wl("");
                        }
                        else
                        {
                            ConsoleHelper.ConsoleHelper.wl("");
                            ConsoleHelper.ConsoleHelper.wl("No Orders!");
                            ConsoleHelper.ConsoleHelper.wl("");
                        }
                    }
                    else
                    {
                        ConsoleHelper.ConsoleHelper.wl("Error:DB Cache Directory is NULL! System will be try to create it.", ConsoleColor.Red, ConsoleColor.Black);
                        Directory.CreateDirectory(url);
                    }
                }
                catch (Exception ex)
                {
                    ConsoleHelper.ConsoleHelper.wl("Error:" + ex.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                }
            }
            ConsoleHelper.ConsoleHelper.wl("End Running...");
            boolDBCache = false;
            GC.Collect();
        }
        #endregion
        #region 子过程
        private static string GetUIDfromTransNum(string TransNum)
        {
            string strResult = "";
            string strSQL = "select Applicants from ApplicationInfo where TransNo='" + TransNum + "' ";
            AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(LinkString2);
            DataTable dtSQL = ah.ReturnDataTable(strSQL);
            if (dtSQL != null && dtSQL.Rows.Count > 0)
            {
                strResult = dtSQL.Rows[0]["Applicants"].ToString();
            }
            return strResult;
        }
        private static double GetTotalPricefromUID(string UID)
        {
            double douResult = 0.00;
            double douCPrice = 0.00;
            double douTPrice = 0.00;
            AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(LinkString2);
            string strSQL = "select RestAmount from Users where UID='" + UID + "' and IsDelete=0 ";
            DataTable dtSQL = ah.ReturnDataTable(strSQL);
            if (dtSQL != null && dtSQL.Rows.Count > 0)
            {
                douCPrice = double.Parse(dtSQL.Rows[0]["RestAmount"].ToString());
            }
            else
            {
                douCPrice = 0.00;
            }
            strSQL = "select * from ApplicationInfo where Applicants='" + UID + "' and IsDelete=0 and AppState>=0 and AppState<6 ";
            dtSQL = ah.ReturnDataTable(strSQL);
            if (dtSQL != null && dtSQL.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSQL.Rows)
                {
                    if (dr["TotalPrice"] != null && dr["TotalPrice"].ToString() != "")
                    {
                        douTPrice += double.Parse(dr["TotalPrice"].ToString());
                    }
                }
            }
            else
            {
                douTPrice = 0.00;
            }
            douResult = douCPrice - douTPrice;
            return douResult;
        }
        #endregion
    }
}
