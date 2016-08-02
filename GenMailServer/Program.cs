using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml;

namespace GenMailServer
{
    class Program
    {
        static void Main(string[] args)
        {
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
                    string LinkString1 = xnCon.SelectSingleNode("LinkString1").InnerText;
                    string LinkString2 = xnCon.SelectSingleNode("LinkString2").InnerText;
                    boolstatus = true;                
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

            }
        }
    }
}
