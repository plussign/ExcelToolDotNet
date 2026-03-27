using ExcelTool;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

class Log
{
    public static void WriteLine(string format, params object[] args)
    {
        if (Program.outputLog)
        {
            Console.WriteLine(string.Format(format, args));
        }
    }

    public static void Write(string format, params object[] args)
    {
        if (Program.outputLog)
        {
            Console.Write(string.Format(format, args));
        }
    }
}
