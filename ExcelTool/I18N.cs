using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using libxl;
using ExcelTool.ExcelTool;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace ExcelTool
{
    public static class I18N
    {
        private static List<string> orginalTextList = new List<string>();
        private static Dictionary<string, uint> orginalTextDict = new Dictionary<string, uint>();

        private static Dictionary<string, string> languageTableMap = new Dictionary<string, string>();

        private static HashSet<char> SDFFontCharSet = new HashSet<char>();

        static I18N()
        {
            languageTableMap.Add("简体", "CN");
            languageTableMap.Add("繁体", "TW");
            languageTableMap.Add("英文", "EN");
            languageTableMap.Add("日语", "JP");
            languageTableMap.Add("韩文", "KR");
            languageTableMap.Add("泰文", "TH");
            languageTableMap.Add("越南", "VN");
        }

        private static bool hasChinese(string text)
        {
            return Regex.IsMatch(text, @"[\u4e00-\u9fa5]");
        }

        private static bool hasChinese(char c)
        {
            return Regex.IsMatch(c.ToString(), @"[\u4e00-\u9fa5]");
        }

        public static uint RegisterText(string text, bool ChineseOnly)
        {
            if (string.IsNullOrWhiteSpace(text) || (ChineseOnly && !hasChinese(text)))
            {
                return 0;
            }

            text = text.Replace("\r\n", "\n").Replace("\r", "\n");

            if (orginalTextDict.TryGetValue(text, out uint index))
            {
                //已存在的词条
                return index + 1;
            }
            else
            {
                //新添加词条
                uint currentIndex = (uint)orginalTextList.Count;
                orginalTextList.Add(text);
                orginalTextDict.Add(text, currentIndex);

                return currentIndex + 1;
            }
        }

        public static void RegisterSDFText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            for (int index = 0; index < text.Length; index++)
            {
                char currentChar = text[index];

                if (!SDFFontCharSet.Contains(currentChar) && hasChinese(currentChar))
                {
                    SDFFontCharSet.Add(currentChar);
                }
            }
        }

        //Lua字符串表
        private static void writeLuaLanguageTable(string fileName, List<string> textList)
        {
            StringBuilder builder = new StringBuilder();
            if (Program.isDynamicOutPut)
            {
                builder.Append("LanguageDynamicTable=\r\n{\r\n");              
            }
            else
            {
                builder.Append("LanguageTable=\r\n{\r\n");
            }
       
            for (int i = 0; i < textList.Count; ++i)
            {
                builder.AppendFormat("\t{0},\t-- {1}\r\n", Assist.ToLuaStr(textList[i]), i + 1);
            }
            builder.Append("}\r\n");

            DirectoryInfo languageTablePath = new DirectoryInfo("languageTables");
            if (languageTablePath.Exists)
            {
                string filePath = Path.Combine(languageTablePath.FullName, fileName);
                BaseHelper.WriteText(filePath, builder.ToString());
            }
        }

        private static void writeScriptableObjectLanguageTable(List<string> textList)
        {
            StringBuilder builder = new StringBuilder();

            string yamlHeader = @"%YAML 1.1
%TAG !u! tag:unity3d.com,2011:
--- !u!114 &11400000
MonoBehaviour:
  m_ObjectHideFlags: 0
  m_CorrespondingSourceObject: {fileID: 0}
  m_PrefabInstance: {fileID: 0}
  m_PrefabAsset: {fileID: 0}
  m_GameObject: {fileID: 0}
  m_Enabled: 1
  m_EditorHideFlags: 0
  m_Script: {fileID: 11500000, guid: $GUID, type: 3}
  m_Name: $TABLE_NAME
  m_EditorClassIdentifier: 
  defaultLanguage:
";
            const string tableName = "StringTable";
            yamlHeader = yamlHeader.Replace("$TABLE_NAME", tableName);
            string csMetaPath = "../Client/Assets/Scripts/DesignersTable/" + tableName + ".cs.meta";
            /*if (!File.Exists(csMetaPath))
            {
                csMetaPath = csMetaPath.Replace("NekoIDWeb", "Char");
            }
            if (!File.Exists(csMetaPath))
            {
                csMetaPath = csMetaPath.Replace("Char", "LiveConcert");
            }*/

            if (File.Exists(csMetaPath))
            {
                var metaText = File.ReadAllText(csMetaPath);

                Match guidMatch = Regex.Match(metaText, @"guid:\s([a-f0-9]+)");
                string guid = guidMatch.Value.Substring(6);

                yamlHeader = yamlHeader.Replace("$GUID", guid);

                builder.Append(yamlHeader);

                Log.WriteLine("[{0}] GUID: {1}", Path.GetFullPath(csMetaPath), guid);
            }
            else
            {
                Log.WriteLine("[{0}].cs.meta NOT FOUND", tableName);
                return;
            }

            for (int i = 0; i < textList.Count; ++i)
            {
                string s = textList[i];

                if (s.Contains("\"") 
                    || s.Contains("\n") 
                    || s.Contains(":") 
                    || s.Contains("#") 
                    || s.Contains("|") 
                    || s.Contains(">") 
                    || s.Contains("[") 
                    || s.Contains("]") 
                    || s.Contains("{") 
                    || s.Contains("}"))
                {
                    s = s.Replace("\n", "\\n").Replace("\"", "\\\"");
                    s = "\"" + s + "\"";
                }

                builder.AppendLine("  - " + s);
            }

            string scriptableObjectFilename = string.Format("output_asset/{0}.asset", tableName);
            BaseHelper.WriteText(scriptableObjectFilename, builder.ToString());
        }

        //UILayout翻译数据
        private static void writeUILayoutTable(string fileName, Dictionary<string, string> textDict)
        {
            StringBuilder builder = new StringBuilder();
            foreach (var pair in textDict)
            {
                builder.AppendFormat("{0}\b{1}\f", pair.Key, pair.Value);
            }

            DirectoryInfo languageTablePath = new DirectoryInfo("languageTables");
            if (languageTablePath.Exists)
            {
                string filePath = Path.Combine(languageTablePath.FullName, fileName);
                BaseHelper.WriteText(filePath, builder.ToString());
            }
        }

        //输出DSF字体生成需要的字符文件
        private static void writeCharFileForDSFFontGeneration(string fileName, HashSet<char> charSet)
        {
            StringBuilder builder = new StringBuilder();

            builder.Append("0123456789！@#￥%……&*（）——+。，《》？、；：‘“”【】{}/*-+~!@#$%^&*()_+`abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ");
            foreach (var c in charSet)
            {
                builder.Append(c);
            }

            DirectoryInfo languageTablePath = new DirectoryInfo("output_textmeshpro_text");
            if (languageTablePath.Exists)
            {
                string filePath = Path.Combine(languageTablePath.FullName, fileName);
                BaseHelper.WriteText(filePath, builder.ToString());
            }
        }

        private static bool tryGetCellString(Sheet sheet, int row, int col, out string value)
        {
            var cellType = sheet.cellType(row, col);
            if (cellType == libxl.CellType.CELLTYPE_BLANK || cellType == libxl.CellType.CELLTYPE_EMPTY)
            {
                value = string.Empty;
                return false;
            }
            else
            {
                if (cellType == libxl.CellType.CELLTYPE_STRING)
                {
                    value = sheet.readStr(row, col);
                }
                else if (cellType == libxl.CellType.CELLTYPE_NUMBER)
                {
                    value = sheet.readNum(row, col).ToString();
                }
                else
                {
                    value = sheet.readStr(row, col);
                }

                return true;
            }
        }

        public static bool WriteLanguageTables()
        {
            DirectoryInfo languageTablePath = new DirectoryInfo("languageTables");
            if (!languageTablePath.Exists)
            {
                languageTablePath.Create();
            }

            //写入原始语言表
            //writeLuaLanguageTable("lang_default.bytes", orginalTextList);
            writeLuaLanguageTable("lang_default.lua", orginalTextList);
            writeScriptableObjectLanguageTable(orginalTextList);

            writeCharFileForDSFFontGeneration("TextMeshPro.txt", SDFFontCharSet);

            List<string> layoutTextList = loadUILayoutText();

            DirectoryInfo i18nPath = new DirectoryInfo("i18n");
            if (i18nPath.Exists)
            {
                FileInfo[] i18nTables = i18nPath.GetFiles("*.xls");

                foreach (FileInfo fileInfo in i18nTables)
                {
                    string translatedFileName = fileInfo.Name;
                    string languageCode = string.Empty;

                    foreach (var pair in languageTableMap)
                    {
                        if (translatedFileName.Contains(pair.Key))
                        {
                            languageCode = pair.Value;
                            break;
                        }
                    }

                    if (string.IsNullOrEmpty(languageCode))
                    {
                        Log.WriteLine("翻译文件[{0}]没有映射到国际化语言表格 languageTableMap", translatedFileName);
                        continue;
                    }
                    
                    //Lua字符串表生成
                    string languageTableName = string.Format("lang_{0}.lua", languageCode);
                    Log.WriteLine("从翻译文件[{0}]生成国际化语言表格[{1}]", translatedFileName, languageTableName);
                    
                    Dictionary<string, string> translatedText = new Dictionary<string, string>(150000);

                    #region Excel文件读取处理部分
                    Book book = XlsLoader.LoadBook(fileInfo.FullName);
                    if (book == null)
                    {
                        continue;
                    }

                    Sheet sheet = book.getSheet(0);

                    int iRow = 2;
                    while (true)
                    {
                        string orginal;
                        bool hasNext = tryGetCellString(sheet, iRow - 1, 0, out orginal);

                        if (!hasNext)
                        {
                            break;
                        }

                        orginal = orginal.Replace("\r\n", "\n").Replace("\r", "\n");

                        if (!translatedText.ContainsKey(orginal))
                        {
                            string translated = sheet.readStr(iRow - 1, 1);
                            translatedText.Add(orginal, translated);
                        }

                        ++iRow;
                    }

                    book.Dispose();
                    book = null;
                    #endregion

                    List<string> translatedTextList = new List<string>();
                    for (int i = 0; i < orginalTextList.Count; ++i)
                    {
                        string orginalText = orginalTextList[i];
                        if (!translatedText.TryGetValue(orginalText, out string translated))
                        {
                            //Log.WriteLine("词条[{0}]没有译文，使用原词条", orginalText);
                            translated = orginalText;
                        }
                        else if (string.IsNullOrWhiteSpace(translated))
                        {
                            //Log.WriteLine("词条[{0}]译文为空，使用原词条", orginalText);
                            translated = orginalText;
                        }
                        translatedTextList.Add(translated);
                    }

                    //写入Lua翻译语言表
                    writeLuaLanguageTable(languageTableName, translatedTextList);

                    //UILayout翻译数据生成
                    string uilayoutTableName = string.Format("ui_{0}.lang", languageCode);
                    Log.WriteLine("从翻译文件[{0}]生成国际化UILayout翻译数据[{1}]", translatedFileName, uilayoutTableName);

                    Dictionary<string, string> layoutTextDict = new Dictionary<string, string>();
                    for (int i = 0; i < layoutTextList.Count; ++i)
                    {
                        string orginalText = layoutTextList[i];
                        if (!translatedText.TryGetValue(orginalText, out string translated))
                        {
                            //Log.WriteLine("词条[{0}]没有译文，使用原词条", orginalText);
                            translated = orginalText;
                        }
                        else if (string.IsNullOrWhiteSpace(translated))
                        {
                            //Log.WriteLine("词条[{0}]译文为空，使用原词条", orginalText);
                            translated = orginalText;
                        }
                        layoutTextDict.Add(orginalText, translated);
                    }

                    //写入UILayout翻译语言表
                    writeUILayoutTable(uilayoutTableName, layoutTextDict);
                }
            }

            return true;
        }

        //UI内字符串
        private static List<string> loadUILayoutText()
        {
            List<string> layoutTextList = new List<string>();
            DirectoryInfo i18nPath = new DirectoryInfo("i18n");

            if (!i18nPath.Exists)
            {
                i18nPath.Create();
                return layoutTextList;
            }

            FileInfo fileLayoutText = new FileInfo(Path.Combine(i18nPath.FullName, "layoutText.txt"));
            if (!fileLayoutText.Exists)
            {
                return layoutTextList;
            }

            StreamReader sr = new StreamReader(fileLayoutText.FullName, Encoding.UTF8);
            StringBuilder sb = new StringBuilder();
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                sb.Append(line);
                sb.Append('\n');
            }
            string layoutText = sb.ToString();
            string [] layoutTextArray = layoutText.Split('\b');

            layoutTextList.AddRange(layoutTextArray);

            return layoutTextList;
        }

        //更新译文文件，写入尚未翻译的词条
        private static int processLanguageExcel(FileInfo excelFile)
        {
            Book book = XlsLoader.LoadBook(excelFile.FullName);
            if (book == null)
            {
                return -1;
            }

            Sheet sheet = book.getSheet(0);

            Dictionary<string, int> translatedText = new Dictionary<string, int>();

            int iRow = 2;
            while (true)
            {
                string text;
                bool hasNext = tryGetCellString(sheet, iRow - 1, 0, out text);

                if (!hasNext)
                {
                    break;
                }

                text = text.Replace("\r\n", "\n").Replace("\r", "\n");

                if (!translatedText.ContainsKey(text))
                {
                    translatedText.Add(text, 1);
                }
                /*else
                {
                    Log.WriteLine("词条[{0}]重复", text);
                }*/

                ++iRow;
            }

            sheet = null;

            book.Dispose();
            book = null;

            Log.WriteLine("[{0}]个词条已经翻译", translatedText.Count);

            List<string> untranslatedText = new List<string>();
            for (int i = 0; i < orginalTextList.Count; ++i)
            {
                string currentText = orginalTextList[i];
                if (!translatedText.ContainsKey(currentText))
                {
                    untranslatedText.Add(currentText);
                }
            }

            int newRows = untranslatedText.Count;
            Log.WriteLine("新增[{0}]个词条", newRows);

            if (newRows <= 0)
            {
                return 0;
            }

            try
            {
                IWorkbook workbook;
                using (FileStream fs = new FileStream(excelFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    workbook = WorkbookFactory.Create(fs);
                }

                ISheet worksheet = workbook.GetSheetAt(0);
                if (worksheet == null)
                {
                    return -1;
                }

                for (int r = 0; r < newRows; ++r)
                {
                    int rowIndex = iRow - 1 + r; // iRow is 1-based from the reading loop
                    IRow row = worksheet.GetRow(rowIndex);
                    if (row == null) row = worksheet.CreateRow(rowIndex);
                    ICell cell = row.GetCell(0);
                    if (cell == null) cell = row.CreateCell(0);
                    cell.SetCellValue(untranslatedText[r]);
                }

                using (FileStream fs = new FileStream(excelFile.FullName, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }

                workbook.Close();
                workbook = null;
            }
            catch(System.Exception e)
            {
                Log.WriteLine("沒有写入成功 原因[{0}]",e);
            }

            return newRows;
        }

        //更新译文文件，写入尚未翻译的词条
        public static void i18nSync()
        {
            DirectoryInfo i18nPath = new DirectoryInfo("i18n");
            if (!i18nPath.Exists)
            {
                i18nPath.Create();
                return;
            }
            else
            {
                //将Layout中的字符串加入待翻译表
                List<string> layoutTextList = loadUILayoutText();
                for (int i = 0; i < layoutTextList.Count; ++i)
                {
                    RegisterText(layoutTextList[i], true);
                }

                Log.WriteLine("共计[{0}]个需要国际化的词条", orginalTextList.Count);

                FileInfo[] i18nTables = i18nPath.GetFiles("*.xls");

                foreach (FileInfo fileInfo in i18nTables)
                {
                    fileInfo.Attributes = FileAttributes.Normal;

                    Log.WriteLine("正在更新国际化文本表[{0}]", fileInfo.Name);

                    processLanguageExcel(fileInfo);
                }
            }
        }
    }
}
