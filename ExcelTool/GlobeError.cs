using System;
using System.Collections.Generic;

namespace ExcelTool
{
    class GlobeInfo
    {
        static public List<string> infoList = new List<string>();

        static public void Push(string str)
        {
            infoList.Add(str);
        }

        static public void Report()
        {
            if (infoList.Count > 0)
            {
                Console.WriteLine("---------信息----------");
                foreach (string v in infoList)
                {
                    Console.WriteLine(v);
                }
                Console.WriteLine("---------信息----------");
                infoList.Clear();
            }
        }
    }

    class GlobeWarning
    {
        static public List<string> warningList = new List<string>();

        static public void Push(string str)
        {
            warningList.Add(str);
        }

        static public void Report()
        {
            if (!Program.outputLog || warningList.Count <= 0)
            {
                return;
            }

            Console.WriteLine("---------警告----------");
            foreach (string v in warningList)
            {
                Console.WriteLine(v);
            }
            Console.WriteLine("---------警告----------");
            warningList.Clear();
        }
    }

    class GlobeError
    {
        static public List<string> errList = new List<string>();

        static public void Push(string str)
        {
            errList.Add(str);
        }

        static public bool Report()
        {
            if (errList.Count > 0)
            {
                Console.WriteLine("---------错误----------");
                foreach (string v in errList)
                {
                    Console.WriteLine(v);
                }
                Console.WriteLine("---------错误----------");
                errList.Clear();
                return true;
            }
            else
            {
                Console.WriteLine("处理完毕");
                return false;
            }
        }
    }
}
