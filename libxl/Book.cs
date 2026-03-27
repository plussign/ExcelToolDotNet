using System;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;

namespace libxl
{
    public class Book : IDisposable
    {
        protected IWorkbook workbook;

        public void Dispose()
        {
            if (workbook != null)
            {
                workbook.Close();
                workbook = null;
            }
            GC.SuppressFinalize(this);
        }

        ~Book()
        {
            Dispose();
        }

        public bool load(string filename)
        {
            try
            {
                using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    workbook = WorkbookFactory.Create(fs);
                }
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void save(string filename)
        {
            using (FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }

        public Sheet getSheet(int index)
        {
            ISheet s = workbook.GetSheetAt(index);
            if (s == null)
            {
                return null;
            }
            return new Sheet(s);
        }

        public int sheetCount()
        {
            return workbook.NumberOfSheets;
        }

        public IWorkbook GetNPOIWorkbook()
        {
            return workbook;
        }
    }

    public class BinBook : Book
    {
        public BinBook()
        {
            // Workbook will be created on load; default to HSSFWorkbook for .xls
        }
    }

    public class XmlBook : Book
    {
        public XmlBook()
        {
            // Workbook will be created on load; default to XSSFWorkbook for .xlsx
        }
    }
}
