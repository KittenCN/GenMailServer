using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleHelper
{
    public class ConsoleHelper
    {
        public static ConsoleColor ccDefColor = ConsoleColor.Green;
        public static ConsoleColor ccDefBackColer = ConsoleColor.Black;
        public static void wl(string strValues)
        {
            Console.WriteLine(strValues);
        }
        public static void wl(string strValues,ConsoleColor cc)
        {
            Console.ForegroundColor = cc;
            Console.WriteLine(strValues);
            Console.ForegroundColor = ccDefColor;
        }
        public static void wl(string strValues, ConsoleColor cc,ConsoleColor bcc)
        {
            Console.ForegroundColor = cc;
            Console.BackgroundColor = bcc;
            Console.WriteLine(strValues);
            Console.ForegroundColor = ccDefColor;
            Console.BackgroundColor = ccDefBackColer;
        }
        public static void wrr(string strValues)
        {
            Console.Write("\r" + strValues);
        }
        public  static void wrr(string strValues,ConsoleColor cc, ConsoleColor bcc)
        {
            Console.ForegroundColor = cc;
            Console.BackgroundColor = bcc;
            Console.Write("\r" + strValues);
            Console.ForegroundColor = ccDefColor;
            Console.BackgroundColor = ccDefBackColer;
        }
    }
}
