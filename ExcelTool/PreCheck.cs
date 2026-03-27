using ExcelTool.ExcelTool;
using libxl;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelTool
{
    public partial class ConvertTool
    {
        private InputConfig preCheckInput;
        private OutputConfig preCheckOutput;
        private FieldConfig preCheckFieldConfig;
        private List<string> preCheckPrimarys;
        private int preCheckNum = 0;
        // 预加载检测配置
        public bool PreCheckLoad(string configName)
        {
            if (!File.Exists(Path.Combine("config", configName)))
            {
                GlobeError.Push("配置文件不存在 " + configName);
                return false;
            }

            XmlDocument doc = new XmlDocument();
            try
            {
                string xmlContent = File.ReadAllText(Path.Combine("config", configName), Encoding.UTF8);
                doc.LoadXml(xmlContent);
            }
            catch (System.Exception e)
            {
                GlobeError.Push("配置文件加载失败= " + configName+ "," + e.ToString());
                return false;
            }

            if (doc.DocumentElement == null)
            {
                GlobeError.Push("配置文件载入失败" + configName);
                return false;
            }


            foreach (XmlElement child in doc.DocumentElement)
            {
                if (!LoadPreCheckTable(child, configName))
                {
                    return false;
                }
            }
            return true;
        }

        private bool LoadPreCheckTable(XmlElement root, string configName)
        {
            preCheckInput = new InputConfig(Program.special_channel);
            preCheckInput.Load(root);

            preCheckOutput = new OutputConfig();
            preCheckOutput.Load(root);

            preCheckFieldConfig = new FieldConfig();
            preCheckFieldConfig.Load(root, configName);

            preCheckPrimarys = new List<string>();
            string csvContent = string.Empty;

            return LoadPreCheckFile(configName.Substring(0, configName.IndexOf(".xml")));
        }

        private bool LoadPreCheckFile(string strCfgXmlKey)
        {
            foreach (InputConfig.SourceFileInfo fileInfo in preCheckInput.files)
            {
                string file = fileInfo.fileName;
                SheetCache sheet = SheetCacheMgr.GetCache(file);
                if (sheet == null)
                {
                    Log.Write("==>>Open[{0}]...", file);

                    string postfix = Path.GetExtension(file).ToLower();
                    string oldDirname = Path.GetDirectoryName(file);
                    if (!File.Exists(file))
                    {
                        if (postfix == ".xls")
                        {
                            string plainFileName = Path.Combine(oldDirname, Path.GetFileNameWithoutExtension(file));
                            Console.WriteLine($"\n文件[{plainFileName}]不存在，尝试替换后缀名后重新寻找...");
                            if (File.Exists(plainFileName + ".xlsx"))
                            {
                                file = plainFileName + ".xlsx";
                            }
                        }
                        else if (postfix == ".xlsx")
                        {
                            string plainFileName = Path.Combine(oldDirname, Path.GetFileNameWithoutExtension(file));
                            Console.WriteLine($"\n文件[{plainFileName}]不存在，尝试替换后缀名后重新寻找...");
                            if (File.Exists(plainFileName + ".xls"))
                            {
                                file = plainFileName + ".xls";
                            }
                        }
                    }
                    
                    if (!File.Exists(file))
                    {
                        GlobeError.Push(string.Format("文件不存在:{0}", file));
                        return false;
                    }

                    Book book = XlsLoader.LoadBook(file);
                    if (book == null)
                    {
                        return false;
                    }

                    sheet = new SheetCache(book.getSheet(0));

                    book.Dispose();
                    book = null;

                    SheetCacheMgr.AddExcelFileCache(file, sheet);

                    Log.WriteLine("Done");
                }
                else
                {
                    Log.WriteLine("$$>>ReUse[{0}]...", file);
                }

                if (sheet == null)
                {
                    GlobeError.Push(string.Format("无法获得工作薄:{0}", file));
                    return false;
                }
                if (!preCheckFieldConfig.LoadSlotInfo(sheet, file))
                {
                    return false;
                }

                Dictionary<string, Dictionary<int, string>> dicColumns = new Dictionary<string, Dictionary<int, string>>();
                for (int i = 1; i < sheet.lastRow(); ++i)
                {
                    if (!allLoadCfgData.ContainsKey(strCfgXmlKey))
                    {
                        allLoadCfgData[strCfgXmlKey] = dicColumns;
                    }

                    ReadPreCheckCsvLine(file, sheet, i, strCfgXmlKey, allLoadCfgData[strCfgXmlKey]);
                }
                
                GC.Collect();
            }
            
            return true;
        }

        public string ReadPreCheckLine(List<string> input)
        {
            StringBuilder content = new StringBuilder();
            string key = string.Empty;
            for (int i = 0; i < preCheckFieldConfig.excelFields.Count; ++i)
            {
                var field = preCheckFieldConfig.excelFields[i];
                string cell = string.Empty;
                if (field.mType.Equals("string"))
                {
                    cell = Assist.ToLuaStr(input[i]);
                }
                else
                {
                    cell = input[i];
                }

                if (content.Length > 0)
                {
                    content.Append(",");
                }

                content.Append(cell);
                if (field.isPrimary)
                {
                    key = cell;
                }
            }

            return string.Format("[{0}]={{{1}}},\r\n", key, content.ToString());
        }

        public void ReadPreCheckCsvLine(string filename, SheetCache sheet, int line, string strCfgXmlKey, Dictionary<string, Dictionary<int, string>> dicColumns)
        {
            for (int i = 0; i < preCheckFieldConfig.excelFields.Count; ++i)
            {

                ExcelField field = preCheckFieldConfig.excelFields[i];
                CellType ct = sheet.cellType(line, field.srcSlot);
                string s = null;
                if (field.mType.Equals("double"))
                {
                    //浮点数型
                    if (ct == CellType.CELLTYPE_BLANK || ct == CellType.CELLTYPE_EMPTY)
                    {
                        s = "0";
                    }
                    else if (ct == CellType.CELLTYPE_STRING)
                    {
                        string str = sheet.readStr(line, field.srcSlot);
                        double d = 0.0f;
                        if (double.TryParse(str, out d))
                        {
                            s = d.ToString("G");
                        }
                        else
                        {
                            GlobeError.Push(string.Format(
                                "解析{0}单元格错误 行:{1}, 列:{2}, \n约束格式为:浮点数, 输入格式为:字符串, 输入值为: {3}, 无法将该值转换为浮点数",
                                filename, line + 1, field.srcSlot + 1, str));
                            return;
                        }
                    }
                    else
                    {
                        double d = sheet.readNum(line, field.srcSlot);
                        s = d.ToString("G");
                    }
                }
                else if (field.mType.Equals("int"))
                {
                    //整数型
                    int d = 0;
                    if (ct == CellType.CELLTYPE_BLANK || ct == CellType.CELLTYPE_EMPTY)
                    {
                        s = "0";
                    }
                    else if (ct == CellType.CELLTYPE_STRING)
                    {
                        string str = sheet.readStr(line, field.srcSlot);
                        if (int.TryParse(str, out d))
                        {
                            s = d.ToString();
                        }
                        else
                        {
                            GlobeError.Push(string.Format("解析{0}单元格错误 行:{1}, 列:{2}, \n约束格式为:整数" +
                            ", 输入格式为:字符串, 输入值为: {3}, 无法将该值转换为整数",
                                filename, line + 1, field.srcSlot + 1, str));
                            return;
                        }
                    }
                    else
                    {
                        d = (int)sheet.readNum(line, field.srcSlot);
                        s = d.ToString();
                    }
                    //allLoadCfgData
                }
                else if (field.mType.Equals("centimeter") || field.mType.Equals("decimeter") || field.mType.Equals("ratio") || field.mType.Equals("millimetre"))
                {
                    //整数型
                    int d = 0;
                    if (ct == CellType.CELLTYPE_BLANK || ct == CellType.CELLTYPE_EMPTY)
                    {
                        s = "0";
                    }
                    else if (ct == CellType.CELLTYPE_STRING)
                    {
                        string str = sheet.readStr(line, field.srcSlot);
                        if (int.TryParse(str, out d))
                        {
                            s = d.ToString();
                        }
                        else
                        {
                            GlobeError.Push(string.Format("解析{0}单元格错误 行:{1}, 列:{2}, \n约束格式为:整数" +
                            ", 输入格式为:字符串, 输入值为: {3}, 无法将该值转换为整数",
                                filename, line + 1, field.srcSlot + 1, str));
                            return;
                        }
                    }
                    else
                    {
                        d = (int)sheet.readNum(line, field.srcSlot);
                        s = d.ToString();
                    }
                    CompareNumIsTrue(field, line, filename, d);

                    //allLoadCfgData
                }
                else if (field.mType.Equals("string"))
                {
                    //字符型
                    if (ct == CellType.CELLTYPE_NUMBER)
                    {
                        s = sheet.readNum(line, field.srcSlot).ToString();
                    }
                    else if (ct == CellType.CELLTYPE_STRING)
                    {
                        s = sheet.readStr(line, field.srcSlot);
                        if (s == null)
                        {
                            s = "";
                        }
                    }
                    else if (ct == CellType.CELLTYPE_BLANK || ct == CellType.CELLTYPE_EMPTY)
                    {
                        s = "";
                    }
                    else
                    {
                        GlobeError.Push(string.Format("解析{0}错误, 行:{1}, 列:{2}, 单元格数据错误",
                            filename, line + 1, field.srcSlot + 1));
                        return;
                    }
                }
                else
                {
                    //枚举值
                    string enumKey = sheet.readStr(line, field.srcSlot);

                    if (enumKey == null)
                    {
                        string cet = sheet.cellType(line, field.srcSlot).ToString();

                        GlobeError.Push(string.Format("解析{0}错误, 无法读取excel文件, 行:{1}, 列:{2}, CellType={3}",
                            filename, line + 1, field.srcSlot + 1, cet));
                        return;
                    }

                    s = enumList.GetEnumValue(field.mType, enumKey);

                    if (s == null)
                    {
                        GlobeError.Push(string.Format("解析{0}错误, 行:{1}, 列:{2}, 键值:{3}, 类型={4}, 枚举={5}, 枚举读取失败",
                         filename, line + 1, field.srcSlot + 1, field.key, field.mType, enumKey));
                        return;
                    }
                }

                if (s != null)
                {
                    if (allLoadCfgData.ContainsKey(strCfgXmlKey))
                    {
                        tmpCheckList = null;
                        //Log.WriteLine(strCfgXmlKey + " 列名字:" + field.key + " dicColumn:cocunt:" + dicColumns.Count);
                        if (dicColumns.ContainsKey(field.key))
                        {
                            dicColumns.TryGetValue(field.key, out Dictionary<int, string> tmpList);
                            tmpList[line] = s;
                            dicColumns[field.key] = tmpList;
                            //Log.WriteLine(strFileKey + " 已经拥有 " + field.name + " 值:" + s + " 该列个数:" + tmpList.Count);
                        }
                        else
                        {
                            //Log.WriteLine(strFileKey + "  没有这个列:" + field.name + " 值:" + preCheckNum);
                            tmpCheckList = new Dictionary<int, string>
                            {
                                [line] = s
                            };
                            dicColumns.Add(field.key, tmpCheckList);
                        }
                    }

                }
            }
            }
        


    }
}
