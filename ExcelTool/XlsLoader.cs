using libxl;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.ExcelTool
{
    class XlsLoader
    {
		/// 装备表@测试.xls, 专门给测试用的数据。
        const string testFilenameTail = "@测试";

        static public Book LoadBook(string filename)
        {
            Book book = new Book();

            if (Program.useTestData)
            {
                string newFilename = Path.GetFileNameWithoutExtension(filename) + 
                    testFilenameTail + Path.GetExtension(filename);

                string oldDirname = Path.GetDirectoryName(filename);

                string newPathname = Path.Combine(oldDirname, newFilename);
                if (File.Exists(newPathname))
                {
                    GlobeInfo.Push("[test] 使用测试表格: " + newPathname);
                    filename = newPathname;
                }
            }

            if (!book.load(filename))
            {
                GlobeError.Push(string.Format("无法载入excel文件:{0}", filename));
                return null;
            }

            return book;
        }
    }
}
