using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using libxl;
using ExcelTool.ExcelTool;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;

namespace ExcelTool
{
    public static class I18N
    {
        private static List<string> originalTextList = new List<string>();
        private static Dictionary<string, uint> originalTextDict = new Dictionary<string, uint>();

        private static Dictionary<string, string> languageTableMap = new Dictionary<string, string>();

        private static HashSet<char> SDFFontCharSet = new HashSet<char>();

        static I18N()
        {
            languageTableMap.Add("简体", "CN");
            languageTableMap.Add("繁体", "TW");
            languageTableMap.Add("英文", "EN");
            languageTableMap.Add("日文", "JP");
            languageTableMap.Add("韩文", "KR");
            //languageTableMap.Add("泰文", "TH");
            //languageTableMap.Add("越南", "VN");
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

            if (originalTextDict.TryGetValue(text, out uint index))
            {
                //已存在的词条
                return index + 1;
            }
            else
            {
                //新添加词条
                uint currentIndex = (uint)originalTextList.Count;
                originalTextList.Add(text);
                originalTextDict.Add(text, currentIndex);

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

                if (!SDFFontCharSet.Contains(currentChar))
                {
                    SDFFontCharSet.Add(currentChar);
                }
            }
        }

        //Lua字符串表
        // private static void writeLuaLanguageTable(string fileName, List<string> textList)
        // {
        //     StringBuilder builder = new StringBuilder();
        //     if (Program.isDynamicOutPut)
        //     {
        //         builder.Append("LanguageDynamicTable=\r\n{\r\n");              
        //     }
        //     else
        //     {
        //         builder.Append("LanguageTable=\r\n{\r\n");
        //     }
       
        //     for (int i = 0; i < textList.Count; ++i)
        //     {
        //         builder.AppendFormat("\t{0},\t-- {1}\r\n", Assist.ToLuaStr(textList[i]), i + 1);
        //     }
        //     builder.Append("}\r\n");

        //     DirectoryInfo languageTablePath = new DirectoryInfo("languageTables");
        //     if (languageTablePath.Exists)
        //     {
        //         string filePath = Path.Combine(languageTablePath.FullName, fileName);
        //         BaseHelper.WriteText(filePath, builder.ToString());
        //     }
        // }

        private static string s_stringTableDefineScriptMeta = string.Empty;
        private static void writeScriptableObjectLanguageTable(string langCode, List<string> textList)
        {
            StringBuilder builder = new StringBuilder();

            const string yamlHeaderTemplate = @"%YAML 1.1
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

            if (string.IsNullOrEmpty(s_stringTableDefineScriptMeta))
            {
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
                    s_stringTableDefineScriptMeta = guidMatch.Value.Substring(6);
                    Log.WriteLine("[{0}] GUID: {1}", Path.GetFullPath(csMetaPath), s_stringTableDefineScriptMeta);
                }
                else
                {
                    Log.WriteLine("[{0}].cs.meta NOT FOUND", tableName);
                    return;
                }
            }

            string yamlHeader = yamlHeaderTemplate.Replace("$GUID", s_stringTableDefineScriptMeta);
            string tableNameWithLang = string.IsNullOrWhiteSpace(langCode) ? tableName : $"{tableName}_{langCode}";
            yamlHeader = yamlHeader.Replace("$TABLE_NAME", tableNameWithLang);
            builder.Append(yamlHeader);


            for (int i = 0; i < textList.Count; ++i)
            {
                string s = textList[i];
                RegisterSDFText(s);

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

            string scriptableObjectFilename = string.Format("output_asset/{0}.asset", tableNameWithLang);
            BaseHelper.WriteText(scriptableObjectFilename, builder.ToString());

            if (string.IsNullOrEmpty(langCode))
            {
                Log.WriteLine("\t写入原文语言表格[{0}]", scriptableObjectFilename);
            }
            else
            {
                Log.WriteLine("\t写入翻译[{0}]语言表格[{1}]]", langCode, scriptableObjectFilename);
            }
        }

        //UILayout翻译数据
        private static void writeUILayoutTable(string fileName, List<UILayoutEntry> entries)
        {
            string json = JsonSerializerHelper.SerializeUILayoutEntries(entries);

            DirectoryInfo languageTablePath = new DirectoryInfo("languageTables");
            if (languageTablePath.Exists)
            {
                string filePath = Path.Combine(languageTablePath.FullName, fileName);
                BaseHelper.WriteText(filePath, json);
            }
        }

        //输出DSF字体生成需要的字符文件
        private static void writeCharFileForDSFFontGeneration(string fileName, HashSet<char> charSet)
        {
            StringBuilder builder = new StringBuilder();
            SortedSet<char> allChars = new SortedSet<char>();

            string basicChars = "0123456789！@#￥%……&*（）——+。，《》？、；：‘“”【】{}/*-+~!@#$%^&*()_+`abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            foreach (var c in basicChars)
            {
                allChars.Add(c);
            }

            // 平假名 U+3041 ~ U+3096
            for (char c = '\u3041'; c <= '\u3096'; c++)
            {
                allChars.Add(c);
            }

            // 片假名 U+30A1 ~ U+30FA
            for (char c = '\u30A1'; c <= '\u30FA'; c++)
            {
                allChars.Add(c);
            }

            foreach (var c in charSet)
            {
                allChars.Add(c);
            }

            foreach (var c in allChars)
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
            //writeLuaLanguageTable("lang_default.bytes", originalTextList);
            //writeLuaLanguageTable("lang_default.lua", originalTextList);
            writeScriptableObjectLanguageTable(string.Empty, originalTextList);

            List<string> layoutTextList = loadUILayoutText();

            DirectoryInfo i18nPath = new DirectoryInfo("i18n");
            if (i18nPath.Exists)
            {
                FileInfo[] i18nTables = i18nPath.GetFiles("*.xlsx");

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
                        Log.WriteLine("翻译文件[{0}]没有映射到 工具内Dictionary languageTableMap", translatedFileName);
                        continue;
                    }
                    
                    //Lua字符串表生成
                    //string languageTableName = string.Format("lang_{0}.lua", languageCode);
                    //Log.WriteLine("从翻译文件[{0}]生成国际化语言表格[{1}]", translatedFileName, languageTableName);
                    
                    Dictionary<string, string> translatedText = new Dictionary<string, string>(60000);

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
                        string original;
                        bool hasNext = tryGetCellString(sheet, iRow - 1, 0, out original);

                        if (!hasNext)
                        {
                            break;
                        }

                        original = original.Replace("\r\n", "\n").Replace("\r", "\n");

                        if (!translatedText.ContainsKey(original))
                        {
                            string translated = sheet.readStr(iRow - 1, 1);
                            translatedText.Add(original, translated);
                        }

                        ++iRow;
                    }

                    book.Dispose();
                    book = null;
                    #endregion

                    List<string> translatedTextList = new List<string>();
                    for (int i = 0; i < originalTextList.Count; ++i)
                    {
                        string originalText = originalTextList[i];
                        if (!translatedText.TryGetValue(originalText, out string translated))
                        {
                            //Log.WriteLine("词条[{0}]没有译文，使用原词条", originalText);
                            translated = originalText;
                        }
                        else if (string.IsNullOrWhiteSpace(translated))
                        {
                            //Log.WriteLine("词条[{0}]译文为空，使用原词条", originalText);
                            translated = originalText;
                        }
                        translatedTextList.Add(translated);
                    }

                    //写入Lua翻译语言表
                    //writeLuaLanguageTable(languageTableName, translatedTextList);

                    //写入ScriptableObject翻译语言表
                    writeScriptableObjectLanguageTable(languageCode, translatedTextList);

                    //UILayout翻译数据生成
                    string uilayoutTableName = string.Format("ui_{0}.json", languageCode);
                    Log.WriteLine("从翻译文件[{0}]生成国际化UILayout翻译数据(JSON)[{1}]", translatedFileName, uilayoutTableName);

                    List<UILayoutEntry> layoutTextListWithTranslation = new List<UILayoutEntry>();
                    for (int i = 0; i < layoutTextList.Count; ++i)
                    {
                        string originalText = layoutTextList[i];
                        if (!translatedText.TryGetValue(originalText, out string translated))
                        {
                            //Log.WriteLine("词条[{0}]没有译文，使用原词条", originalText);
                            translated = originalText;
                        }
                        else if (string.IsNullOrWhiteSpace(translated))
                        {
                            //Log.WriteLine("词条[{0}]译文为空，使用原词条", originalText);
                            translated = originalText;
                        }
                        layoutTextListWithTranslation.Add(new UILayoutEntry { Text = originalText, Translation = translated });
                    }

                    //写入UILayout翻译语言表
                    writeUILayoutTable(uilayoutTableName, layoutTextListWithTranslation);
                }
            }

            writeCharFileForDSFFontGeneration("TextMeshPro.txt", SDFFontCharSet);

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

            FileInfo fileLayoutText = new FileInfo(Path.Combine(i18nPath.FullName, "layoutText.json"));
            if (!fileLayoutText.Exists)
            {
                return layoutTextList;
            }

            string jsonText = File.ReadAllText(fileLayoutText.FullName, Encoding.UTF8);
            var jsonList = JsonSerializerHelper.DeserializeStringList(jsonText);

            if (jsonList != null)
            {
                layoutTextList.AddRange(jsonList);
            }

            return layoutTextList;
        }

        //更新译文文件，写入尚未翻译的词条
        private static int processLanguageExcel(FileInfo excelFile, bool ignoreTranslated)
        {
            HashSet<string> translatedText;
            int iRow = 2;

            if (!ignoreTranslated)
            {
                Book book = XlsLoader.LoadBook(excelFile.FullName);
                if (book == null)
                {
                    return -1;
                }

                Sheet sheet = book.getSheet(0);
                translatedText = new HashSet<string>(sheet.lastRow() + 1);

                while (true)
                {
                    string text;
                    bool hasNext = tryGetCellString(sheet, iRow - 1, 0, out text);

                    if (!hasNext)
                    {
                        break;
                    }

                    text = text.Replace("\r\n", "\n").Replace("\r", "\n");

                    if (!translatedText.Contains(text))
                    {
                        translatedText.Add(text);
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

                Log.WriteLine("\t[{0}]个词条先前已经提取", translatedText.Count);
            }
            else
            {
                translatedText = new HashSet<string>(0);
                Log.WriteLine("\t忽略已提取的词条，直接写入所有待翻译词条");
            }

            List<string> untranslatedText = new List<string>();
            for (int i = 0; i < originalTextList.Count; ++i)
            {
                string currentText = originalTextList[i];
                if (!translatedText.Contains(currentText))
                {
                    untranslatedText.Add(currentText);
                }
            }

            int newRows = untranslatedText.Count;
            Log.WriteLine("\t新增[{0}]个词条", newRows);

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
                Log.WriteLine($"\t沒有写入成功 异常：\n{e}\n");
            }

            return newRows;
        }

        //更新译文文件，写入尚未翻译的词条
        public static void i18nSync(bool ignoreTranslated)
        {
            DirectoryInfo i18nPath = new DirectoryInfo("i18n");
            if (!i18nPath.Exists)
            {
                Log.WriteLine("i18n目录不存在，创建目录[{0}] 请再次运行程序", i18nPath.FullName);
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

                Log.WriteLine($"\n\n共计[{originalTextList.Count}]个需要国际化的词条\n");

                FileInfo[] i18nTables = i18nPath.GetFiles("*.xlsx");

                foreach (FileInfo fileInfo in i18nTables)
                {
                    fileInfo.Attributes = FileAttributes.Normal;

                    Log.WriteLine($"\n正在更新国际化文本表[{fileInfo.Name}]");

                    processLanguageExcel(fileInfo, ignoreTranslated);
                }
            }
        }
    }
}
