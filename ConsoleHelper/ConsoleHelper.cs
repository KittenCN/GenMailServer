using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ConsoleHelper
{
    public class ConsoleHelper
    {
        public static ConsoleColor ccDefColor = ConsoleColor.Green;
        public static ConsoleColor ccDefBackColer = ConsoleColor.Black;
        public static string GenLinkString = "./DB/GenMailServer.accdb";
        public static Boolean boolLog = true;      
        public static void wl(string strValues)
        {
            Console.WriteLine(strValues);
            Log(strValues);
        }
        public static void wl(string strValues,Boolean boolLogFlag)
        {
            Console.WriteLine(strValues);
            if(boolLogFlag==true)
            {
                Log(strValues);
            }
        }
        public static void wl(string strValues,ConsoleColor cc)
        {
            Console.ForegroundColor = cc;
            Console.WriteLine(strValues);
            Console.ForegroundColor = ccDefColor;
            Log(strValues);
        }
        public static void wl(string strValues, ConsoleColor cc,ConsoleColor bcc)
        {
            Console.ForegroundColor = cc;
            Console.BackgroundColor = bcc;
            Console.WriteLine(strValues);
            Console.ForegroundColor = ccDefColor;
            Console.BackgroundColor = ccDefBackColer;
            Log(strValues);
        }
        public static void wrr(string strValues,Boolean boolFlag)
        {
            boolLog = boolFlag;
            Console.Write("\r" + strValues);
            Log(strValues);
            boolLog = true;
        }
        public  static void wrr(string strValues,ConsoleColor cc, ConsoleColor bcc, Boolean boolFlag)
        {
            boolLog = boolFlag;
            Console.ForegroundColor = cc;
            Console.BackgroundColor = bcc;
            Console.Write("\r" + strValues);
            Console.ForegroundColor = ccDefColor;
            Console.BackgroundColor = ccDefBackColer;
            Log(strValues);
            boolLog = true;
        }

        public static void cInitiaze()
        {
            Console.ForegroundColor = ccDefColor;
            Console.BackgroundColor = ccDefBackColer;
            Console.WindowWidth = 120;
            Console.WindowHeight = 33;
            Console.Title = "GMS-General Mail Server";
        }
        public static void Log(string LogBody)
        {
            if(boolLog==true && LogBody != null && LogBody != "")
            {
                LogBody = LogBody.Replace("'", "#");
                LogBody = LogBody.Replace("\"", "#");
                string strDT = DateTime.Now.ToString("yyyyMMdd");
                string strLinkName = "Log" + strDT + ".accdb";
                string strLinkString = "./DB/" + strLinkName;
                if (!File.Exists(strLinkString))
                {
                    File.Copy("./DB/LogTemp.accdb", strLinkString);
                }
                AccessHelper.AccessHelper ah = new AccessHelper.AccessHelper(strLinkString);
                string sql = "insert into Log(Log,LogDateTime) ";
                sql = sql + " values('" + LogBody + "',#" + DateTime.Now.ToString() + "#) ";
                ah.ExecuteNonQuery(sql);
            }
        }
    }
}
