using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    public class BaseHelper
    {
        public static void WriteBin(string filename, byte[] bytes)
        {
            string dir = Path.GetDirectoryName(filename);
            if (!Directory.Exists(dir) && dir.Length > 0)
            {
                Directory.CreateDirectory(dir);
            }

            if (File.Exists(filename))
            {
                FileInfo fi = new FileInfo(filename)
                {
                    Attributes = FileAttributes.Normal
                };
                File.Delete(filename);
            }
            FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write);
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();
        }

        public static void WriteText(string filename, string text)
        {
            string dir = Path.GetDirectoryName(filename);
            if (!Directory.Exists(dir) && dir.Length > 0)
            {
                Directory.CreateDirectory(dir);
            }

            if (File.Exists(filename))
            {
                FileInfo fi = new FileInfo(filename)
                {
                    Attributes = FileAttributes.Normal
                };
                File.Delete(filename);
            }
            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(text);
            FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write);
            byte[] head = new byte[3] { 0xEF, 0xBB, 0xBF };
            fs.Write(head, 0, 3);
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();
        }

        public static byte[] Meger(List<byte[]> input)
        {
            long allLen = 0;
            for (int i = 0; i < input.Count; ++i)
            {
                allLen+= input[i].Length;
            }

            byte[] dst = new byte[allLen];
            long currLen = 0;
            for(int i=0; i<input.Count; ++i)
            {
                input[i].CopyTo(dst, currLen);
                currLen += input[i].Length;
            }

            return dst;
        }

        public static void WriteTextNoBOM(string filename, string text)
        {
            string dir = Path.GetDirectoryName(filename);
            if (!Directory.Exists(dir) && dir.Length > 0)
            {
                Directory.CreateDirectory(dir);
            }

            if (File.Exists(filename))
            {
                FileInfo fi = new FileInfo(filename);
                fi.Attributes = FileAttributes.Normal;
                File.Delete(filename);
            }

            byte[] bytes = System.Text.Encoding.UTF8.GetBytes(text);
            FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write);
            fs.Write(bytes, 0, bytes.Length);
            fs.Close();
        }

    }
}
