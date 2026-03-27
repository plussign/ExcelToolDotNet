using ExcelTool;
using System;

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
