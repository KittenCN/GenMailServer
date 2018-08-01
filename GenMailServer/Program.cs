using System;
using System.IO;
using System.Xml;
using System.Data;
using System.Threading;

namespace GMS
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
        public static Timer tMailMethod;
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
        public static Boolean boolMailFirstRun = false;
        public static Boolean boolDBCheck = false;
        public static DateTime dtCleanLastDay;
        public static Boolean boolCheckMoney = false;
        public static int intDBCPFlag = 0;
        #endregion
        #region Main Method
        static void Main(string[] args)
        {
            ConsoleHelper.ConsoleHelper.cInitiaze();
            ConsoleHelper.ConsoleHelper.wl("Welcome to GMS-General Management Server");
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
            ConsoleHelper.ConsoleHelper.cInitiaze();
            ConsoleHelper.ConsoleHelper.wl("");
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
                    intDBCPFlag = int.Parse(xnCon.SelectSingleNode("DBCacheProcessFlag").InnerText);
                    ConsoleHelper.ConsoleHelper.wl("Reading Config File Successfully...");

                    CheckDB(GenLinkString, GenCheckStr);
                    CheckDB(LinkString1, LinkCheckStr);
                    CheckDB(LinkString2, LinkCheckStr);               

                    intSilentTime = EmailRete;
                    intSecondShow = intMainRate;
                    int3rdShow = intDBCacheRate;
                    //TimerCallback(null);
                    //TimerDBCacheProcess(null);
                    ConsoleHelper.ConsoleHelper.wl("Begin Timer Methods...");
                    tClock = new Timer(TimerClockShow, null, 0, 1000);
                    tMailMethod = new Timer(TimerCallback, null, 0, intMainRate * 1000);
                    do { } while (boolMailFirstRun == false);
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
            if ((DateTime.Now.TimeOfDay.Hours == 0 && DateTime.Now.TimeOfDay.Minutes == 0 && DateTime.Now.TimeOfDay.Seconds < 10 && DateTime.Now.Date != dtCleanLastDay) || intEmailTestFlag == 1)
            {
                AccessHelper.AccessHelper ahLink2 = new AccessHelper.AccessHelper(LinkString2);
                string strSQL = "update SetupConfig set LoginNum = 0";
                if (strSQL != null && strSQL != "")
                {
                    ahLink2.ExecuteNonQuery(strSQL);
                }
                dtCleanLastDay = DateTime.Now.Date;
                ConsoleHelper.ConsoleHelper.wl("");
                ConsoleHelper.ConsoleHelper.wl("Current DateTime is " + DateTime.Now.ToString() + "  ::Run Method of CleanLoginNum!");
            }

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
            if (!boolProcess && !boolDBCache)
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
                boolProcess = true;
                #region DB Check
                ConsoleHelper.ConsoleHelper.wl("");
                ConsoleHelper.ConsoleHelper.wl("Running DB Check Method...");
                //boolClockShow = false;                
                bool boolstatus = false;
                emailHelper.emailHelper eh = new emailHelper.emailHelper();
                try
                {
                    #region Debug Mode
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
                    #endregion
                    if (CheckDB(GenLinkString, GenCheckStr) && CheckDB(LinkString1, LinkCheckStr) && CheckDB(LinkString2, LinkCheckStr))
                    {
                        boolstatus = true;
                        boolDBCheck = true;
                    }
                    else
                    {
                        boolstatus = false;
                        boolDBCheck = false;
                    }
                }
                catch (Exception ex)
                {
                    boolstatus = false;
                    ConsoleHelper.ConsoleHelper.wl("Error:" + ex.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                }
                ConsoleHelper.ConsoleHelper.wl("End DB Check Method Running...");
                //ConsoleHelper.ConsoleHelper.wl("");
                #endregion
                #region Mail Method
                //Mail Method Cancel by 2017.02.13
                if (boolstatus && 1 == 0)
                {
                    //ConsoleHelper.ConsoleHelper.wl("");
                    ConsoleHelper.ConsoleHelper.wl("Running Mail Method...");
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
                    ConsoleHelper.ConsoleHelper.wl("End Mail Method Running...");
                }
                else
                {
                    //Mail Method Cancel by 2017.02.13
                    //ConsoleHelper.ConsoleHelper.wl("Error:" + "Local DB or Remote DB can not be connected!", ConsoleColor.Red, ConsoleColor.Black);
                    ConsoleHelper.ConsoleHelper.wl("Mail Method had been cancelled!");
                }
                #endregion
            }
            ConsoleHelper.ConsoleHelper.wl("Mail Method Running Flag::  boolDBCheck:" + boolDBCheck + "  ||  boolProcess:" + boolProcess + "  ||  boolDBCache:" + boolDBCache);
            ConsoleHelper.ConsoleHelper.wl("");
            boolMailFirstRun = true;
            boolProcess = false;
            GC.Collect();
        }
        #endregion
        #region 数据库及数据表测试
        public static Boolean CheckDB(string strDBAddress, string strTableName)
        {
            Boolean boolResult = false;
            AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(strDBAddress);
            string strConnResult = ah.ConnectTest();
            if (strConnResult == "T")
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
                ConsoleHelper.ConsoleHelper.wl("Error:" + strConnResult.ToString(), ConsoleColor.Red, ConsoleColor.Black);
            }
            return boolResult;
        }
        #endregion
        #region 远程数据操作
        private static void TimerDBCacheProcess(object o)
        {
            if (!boolDBCache && !boolProcess && boolDBCheck)
            {
                boolDBCache = true;
                ConsoleHelper.ConsoleHelper.wl("");
                ConsoleHelper.ConsoleHelper.wl("Running DB Cache Method...");
                #region 总额度计算
                ConsoleHelper.ConsoleHelper.wl("Running Amount Calculation Method...");
                //计算额度
                if (!boolCheckMoney)
                {
                    try
                    {
                        AccessHelper.AccessHelper ahLink2 = new AccessHelper.AccessHelper(LinkString2);
                        boolCheckMoney = true;
                        string strMaxEmpDate = DateTime.Now.AddDays(-180).ToShortDateString();
                        string strBeginDate = DateTime.Now.Year.ToString() + "/2/1";
                        string strEndDate = DateTime.Parse(DateTime.Now.Year.ToString() + "/08/01 00:00:00").AddDays(-1).ToShortDateString();
                        string strBeginDate2 = DateTime.Now.Year.ToString() + "/8/1";
                        string strEndDate2 = DateTime.Parse(DateTime.Now.AddYears(1).Year.ToString() + "/02/01 00:00:00").AddDays(-1).ToShortDateString();


                        string strSQL = "update Users set TotalAmount = 0, UsedAmount = 0, RestAmount = 0 where EmpDate >=#" + strMaxEmpDate + "# or EmpDate is null" ;
                        ahLink2.ExecuteNonQuery(strSQL);
                        if (DateTime.Now.Month == 2 && DateTime.Now.Day == 1)
                        {
                            strSQL = "update Users set UsedAmount = 0, RestAmount = 50000 where EmpDate <#" + strMaxEmpDate + "#";
                            ahLink2.ExecuteNonQuery(strSQL);
                        }
                        if (DateTime.Now.Month >= 2 && DateTime.Now.Month <= 7)
                        {
                            strSQL = "update Users set UsedAmount = 0, RestAmount = 50000 * (datediff('d', EmpDate, #" + strEndDate + "#) / 180) where EmpDate <#" + strMaxEmpDate + "# and EmpDate >= #" + strBeginDate + "# ";
                            ahLink2.ExecuteNonQuery(strSQL);
                        }
                        if (DateTime.Now.Month == 8 && DateTime.Now.Day == 1)
                        {
                            strSQL = "update Users set RestAmount = RestAmount + 50000 where EmpDate <#" + strMaxEmpDate + "#";
                            ahLink2.ExecuteNonQuery(strSQL);
                        }
                        if (DateTime.Now.Month >= 8 || DateTime.Now.Month == 1)
                        {
                            //strSQL = "update Users set UsedAmount = 0, RestAmount = (50000 * (datediff('d', EmpDate, #" + strEndDate + "#) / 180)) + 50000 where EmpDate <#" + strMaxEmpDate + "# and EmpDate < #" + strBeginDate2 + "#";
                            //ahLink2.ExecuteNonQuery(strSQL);
                            strSQL = "update Users set UsedAmount = 0, RestAmount = 50000 * (datediff('d', EmpDate, #" + strEndDate2 + "#) / 180) where EmpDate <#" + strMaxEmpDate + "# and EmpDate >= #" + strBeginDate2 + "#";
                            ahLink2.ExecuteNonQuery(strSQL);
                        }
                        ConsoleHelper.ConsoleHelper.wl("Amount Calculation Running Success!");
                    }
                    catch(Exception ex)
                    {
                        ConsoleHelper.ConsoleHelper.wl("Error:" + ex.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                    }                   
                }
                #endregion
                #region 远程数据缓存处理事件
                ConsoleHelper.ConsoleHelper.wl("Checking DB Cache Directory...");
                int intSuccess = 0;
                int intError = 0;
                string url = LinkString2.Substring(0, LinkString2.LastIndexOf("\\") + 1) + "DBCache\\";
                if (Directory.Exists(url))
                {
                    DirectoryInfo di = new DirectoryInfo(url);
                    ConsoleHelper.ConsoleHelper.wl("Checking DB Cache Files Of Polling...");
                    if(intDBCPFlag == 0)
                    {
                        foreach (FileInfo fi in di.GetFiles("*.accdb"))
                        {
                            var DBCResult = DBCacheProcessCore(fi.FullName, fi.ToString());
                            intSuccess += DBCResult.Item1;
                            intError += DBCResult.Item2;
                        }
                    }
                   else if(intDBCPFlag == 1)
                    {
                        AccessHelper.AccessHelper PurchaseLink = new AccessHelper.AccessHelper(LinkString2);
                        string strSQL = "select UID from [Users] where IsDelete = 0";
                        DataTable dtTable = PurchaseLink.ReturnDataTable(strSQL);
                        if(dtTable.Rows.Count > 0)
                        {
                            foreach(DataRow dr in dtTable.Rows)
                            {
                                if(dr[0].ToString() != "")
                                {
                                    string strName = GetASCII(dr[0].ToString()) + ".accdb";
                                    string strFullName = url + strName;
                                    var DBCResult = DBCacheProcessCore(strFullName, strName);
                                    intSuccess += DBCResult.Item1;
                                    intError += DBCResult.Item2;
                                }
                            }
                        }
                    }
                    else
                    {
                        ConsoleHelper.ConsoleHelper.wl("Error:Config file Error!", ConsoleColor.Red, ConsoleColor.Black);
                    }
                    if (intSuccess > 0 || intError > 0)
                    {
                        ConsoleHelper.ConsoleHelper.wl("Result:");
                        ConsoleHelper.ConsoleHelper.wl("Success:" + intSuccess);
                        ConsoleHelper.ConsoleHelper.wl("Error:" + intError);
                    }
                    else
                    {
                        ConsoleHelper.ConsoleHelper.wl("No Orders!");
                    }
                }
                else
                {
                    ConsoleHelper.ConsoleHelper.wl("Error:DB Cache Directory is NULL! System will be try to create it.", ConsoleColor.Red, ConsoleColor.Black);
                    Directory.CreateDirectory(url);
                }
                ConsoleHelper.ConsoleHelper.wl("End DB Cache Running...");
            }
            ConsoleHelper.ConsoleHelper.wl("DB Cache Running Flag::  boolDBCheck:" + boolDBCheck + "  ||  boolProcess:" + boolProcess + "  ||  boolDBCache:" + boolDBCache);
            ConsoleHelper.ConsoleHelper.wl("");
            boolDBCache = false;
            GC.Collect();
        }
        #endregion
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
        private static double GetTotalPricefromUID(string UID, string TransNum)
        {
            double douResult = 0.00;
            double douCPrice = 0.00;
            double douTPrice = 0.00;
            string strBeginDate = DateTime.Now.Year.ToString() + "/2/1";
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
            strSQL = "select * from ApplicationInfo where Applicants='" + UID + "' and IsDelete=0 and AppState>=0 and AppState<6 and ApplicantsDate >= #" + strBeginDate + "# and TransNo <> '" + TransNum + "'";
            dtSQL = ah.ReturnDataTable(strSQL);
            if (dtSQL != null && dtSQL.Rows.Count > 0)
            {
                foreach (DataRow dr in dtSQL.Rows)
                {
                    //if (dr["TotalPrice"] != null && dr["TotalPrice"].ToString() != "")
                    //{
                    //    douTPrice += double.Parse(dr["TotalPrice"].ToString());
                    //}
                    if(dr["TransNo"] != null && dr["TransNo"].ToString() != "")
                    {
                        string strSQLdetail = "select * from ApplicationDetail where IsDelete = 0 and TransNo='" + dr["TransNo"].ToString() + "' ";
                        AccessHelper.AccessHelper ahdetail = new AccessHelper.AccessHelper(LinkString2);
                        DataTable dtdetail = ahdetail.ReturnDataTable(strSQLdetail);
                        foreach (DataRow drdetail in dtdetail.Rows)
                        {
                            if (drdetail["ItemID"] != null && drdetail["ItemID"].ToString() != "")
                            {
                                if(!GetItemStatus(drdetail["ItemID"].ToString()))
                                {
                                    douTPrice += double.Parse(drdetail["FinalPrice"].ToString());
                                }
                            }
                        }
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
        private static Tuple<int, int> DBCacheProcessCore(string strDB_Add, string strDB_Name)
        {
            int intSuccess = 0;
            int intError = 0;

            try
            {
                AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(strDB_Add);
                string strConnResult = ah.ConnectTest();
                if (strConnResult == "T")
                {
                    ConsoleHelper.ConsoleHelper.wl(Path.GetFileNameWithoutExtension(strDB_Name) + "::DB Cache Main Method...");
                    string strSQL = "select * from AccessQueue";
                    DataTable dtSQL = ah.ReturnDataTable(strSQL);
                    string strLastCtrlID = "";
                    Boolean boolLastCtrlID = true;
                    Boolean boolRepeatCheck = false;
                    foreach (DataRow dr in dtSQL.Rows)
                    {
                        try
                        {                            
                            if (dr["TransNo"] != null && dr["TransNo"].ToString() != "" && dr["operation"].ToString().Substring(0, 5) == "Cache")
                            {
                                string strUID = dr["operation"].ToString().Substring(5);
                                string strDBCheck = dr["SqlStr"].ToString().ToLower().Replace(" ", "").Replace("insertinto", "$");
                                //if (dr["SqlStr"].ToString().ToLower().Substring(1, "create".Length) == "create" && strLastCtrlID != dr["TransNo"].ToString() && strDBCheck.Split('$')[1].ToLower().Substring(1, "applicationInfo".Length) == "applicationInfo")
                                if (dr["SqlStr"].ToString().ToLower().Substring(0, "insert".Length) == "insert")
                                {
                                    if(strDBCheck.Split('$')[1].ToLower().Substring(0, "applicationInfo".Length) == "applicationInfo".ToLower() || strDBCheck.Split('$')[1].ToLower().Substring(0, "[applicationInfo]".Length) == "[applicationInfo]".ToLower())
                                    {
                                        string strSQLin = "select * from ApplicationInfo where TransNo='" + dr["TransNo"].ToString() + "' ";
                                        AccessHelper.AccessHelper ahLink2 = new AccessHelper.AccessHelper(LinkString2);
                                        DataTable dtSQLin = ahLink2.ReturnDataTable(strSQLin);
                                        if (dtSQLin.Rows.Count > 0)
                                        {
                                            strSQLin = "delete from AccessQueue where ID=" + dr["ID"].ToString() + " ";
                                            ah.ExecuteNonQuery(strSQLin);
                                            //strLastCtrlID = dr["TransNo"].ToString();
                                            //boolLastCtrlID = false;
                                            boolRepeatCheck = true;
                                            ConsoleHelper.ConsoleHelper.wl("Order:" + dr["TransNo"].ToString() + "::Repeat Check Fail.Please Check Last Order.", ConsoleColor.Red, ConsoleColor.Black);
                                            intError++;
                                        }
                                        else
                                        {
                                            boolRepeatCheck = false;
                                        }
                                    }
                                    else if((strDBCheck.Split('$')[1].ToLower().Substring(0, "ApplicationDetail".Length) == "ApplicationDetail".ToLower() || strDBCheck.Split('$')[1].ToLower().Substring(0, "[ApplicationDetail]".Length) == "[ApplicationDetail]".ToLower()) && boolRepeatCheck == true)
                                    {
                                        string strSQLin = "select * from ApplicationDetail where TransNo='" + dr["TransNo"].ToString() + "' ";
                                        AccessHelper.AccessHelper ahLink2 = new AccessHelper.AccessHelper(LinkString2);
                                        DataTable dtSQLin = ahLink2.ReturnDataTable(strSQLin);
                                        if (dtSQLin.Rows.Count > 0)
                                        {
                                            strSQLin = "delete from AccessQueue where ID=" + dr["ID"].ToString() + " ";
                                            ah.ExecuteNonQuery(strSQLin);
                                            //strLastCtrlID = dr["TransNo"].ToString();
                                            //boolLastCtrlID = false;
                                            boolRepeatCheck = true;
                                            ConsoleHelper.ConsoleHelper.wl("Order:" + dr["TransNo"].ToString() + "::Repeat Check Fail.Please Check Last Order.", ConsoleColor.Red, ConsoleColor.Black);
                                            intError++;
                                        }
                                        else
                                        {
                                            boolRepeatCheck = false;
                                        }
                                    }
                                    else
                                    {
                                        boolRepeatCheck = false;
                                    }
                                }
                                else
                                {
                                    boolRepeatCheck = false;
                                }
                                if(!boolRepeatCheck)
                                {
                                    if (strLastCtrlID == dr["TransNo"].ToString() && boolLastCtrlID == false)
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
                                        strLastCtrlID = dr["TransNo"].ToString();
                                        boolLastCtrlID = false;
                                        ConsoleHelper.ConsoleHelper.wl("Order:" + dr["TransNo"].ToString() + "::Flag Check Fail.Please Check Last Order.", ConsoleColor.Red, ConsoleColor.Black);
                                        intError++;
                                    }
                                    else
                                    {
                                        strLastCtrlID = dr["TransNo"].ToString();
                                        if (GetTotalPricefromUID(strUID, strLastCtrlID) - double.Parse(dr["Buy"].ToString()) >= 0 || dr["Buy"].ToString() == "" || double.Parse(dr["Buy"].ToString()) == 0 || (strLastCtrlID == dr["TransNo"].ToString() && boolLastCtrlID == true && dr["DetailID"].ToString() == "1"))
                                        {
                                            AccessHelper.AccessHelper ahLink2 = new AccessHelper.AccessHelper(LinkString2);
                                            if (dr["SqlStr"] != null && dr["SqlStr"].ToString() != "")
                                            {
                                                ahLink2.ExecuteNonQuery(dr["SqlStr"].ToString());
                                                string strSQLin = "delete from AccessQueue where ID=" + dr["ID"].ToString() + " ";
                                                ah.ExecuteNonQuery(strSQLin);
                                                boolLastCtrlID = true;
                                                ConsoleHelper.ConsoleHelper.wl("Order:" + dr["TransNo"].ToString() + "::Process Success!");
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
                                            boolLastCtrlID = false;
                                            ConsoleHelper.ConsoleHelper.wl("Order:" + dr["TransNo"].ToString() + "::Quota Check Fail.Please Check It", ConsoleColor.Red, ConsoleColor.Black);
                                            intError++;
                                        }
                                    }
                                }                               
                            }
                            else
                            {
                                ConsoleHelper.ConsoleHelper.wl("TransNo Or Operation Code Check Fail.Please Check It", ConsoleColor.Magenta, ConsoleColor.Black);
                                intError++;
                            }
                        }
                        catch (Exception er)
                        {
                            ConsoleHelper.ConsoleHelper.wl("Error:" + er.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                        }
                    }
                }
                else
                {
                    ConsoleHelper.ConsoleHelper.wl("Error:" + strConnResult.ToString(), ConsoleColor.Red, ConsoleColor.Black);
                }
            }
            catch (Exception ex)
            {
                ConsoleHelper.ConsoleHelper.wl("Error:" + ex.ToString(), ConsoleColor.Red, ConsoleColor.Black);
            }

            return new Tuple<int, int>(intSuccess, intError);
        }
        private static string GetASCII(string strIN)
        {
            string strResult = null;
            byte[] array = System.Text.Encoding.ASCII.GetBytes(strIN);  //数组array为对应的ASCII数组     
            for (int i = 0; i < array.Length; i++)
            {
                int asciicode = (int)(array[i]);
                strResult += Convert.ToString(asciicode);//字符串ASCIIstr2 为对应的ASCII字符串
            }
            return strResult;
        }
        private static Boolean GetItemStatus(string strItemID)
        {
            Boolean boolResult = false;
            string strSQL = "select * from Items where ItemID='" + strItemID + "' and IsDelete = 0";
            AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(LinkString2);
            DataTable dt = ah.ReturnDataTable(strSQL);
            if(dt.Rows[0]["IsSpecial"] != null && dt.Rows[0]["IsSpecial"].ToString() != "")
            {
                if(dt.Rows[0]["IsSpecial"].ToString() == "0")
                {
                    boolResult = false;
                }
                else
                {
                    boolResult = true;
                }
            }
            return boolResult;
        }
        #endregion
    }
}
