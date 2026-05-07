using libxl;
using System;
using System.Collections.Generic;
using System.IO;
using ExcelTool.ExcelTool;

namespace ExcelTool
{
    public enum ReadRawLineResult
    {
        Success,
        Skipped,
        Error
    }

    public class CellDataForLua
    {
        public enum CellTypeForLua
        {
            Standard,
            StringIndex
        }

        public CellTypeForLua type;

        private string standardString;
        private uint stringIndex;

        public CellDataForLua(string str)
        {
            type = CellTypeForLua.Standard;
            standardString = str;
            stringIndex = 0;
        }

        public CellDataForLua(uint index, string str)
        {
            type = CellTypeForLua.StringIndex;
            standardString = str;
            stringIndex = index;
        }
        
        public bool IsBlank
        {
            get
            {
                if (type == CellTypeForLua.Standard)
                {
                    return (string.IsNullOrEmpty(standardString) || standardString.Equals("0"));
                }
                else
                {
                    return false;
                }
            }
        }

        /*
        public override string ToString()
        {
            if (type == CellTypeForLua.Standard)
            {
                return standardString;
            }
            else if (type == CellTypeForLua.StringIndex)
            {
                return string.Format("L[{0}]", stringIndex);
            }
            else
            {
                return "";
            }
        }*/

        public uint GetStringIndex()
        {
            return stringIndex;
        }

        public string GetOriginalString()
        {
            return standardString;
        }

        public string GetSingleRowString()
        {
            return standardString.Replace("\r\n", "\\r\\n").Replace("\r", "\\r").Replace("\n", "\\n");
        }
    }

    public partial class ConvertTool
    {
        private Dictionary<string, string> translatedCsvText;

        private void loadCsvTranslation(string excelFileName)
        {
            DirectoryInfo i18nPath = new DirectoryInfo("i18n");

            if (!i18nPath.Exists)
            {
                i18nPath.Create();
                return;
            }

            string filePath = Path.Combine(i18nPath.FullName, excelFileName);

            Book book = XlsLoader.LoadBook(filePath);
            if (book == null)
            {
                return;
            }

            translatedCsvText = new Dictionary<string, string>();

            Sheet sheet = book.getSheet(0);

            int iRow = 2;
            while (true)
            {
                string original = sheet.readStr(iRow - 1, 0);
                string translated = sheet.readStr(iRow - 1, 1);

                if (string.IsNullOrWhiteSpace(original))
                {
                    break;
                }

                if (!translatedCsvText.ContainsKey(original))
                {
                    translatedCsvText.Add(original, translated);
                }

                ++iRow;
            }

            book.Dispose();
        }

        private bool TryReadPrimaryKeyForDuplicateCheck(
            string filename, SheetCache sheet, int line, ExcelField field, out string value)
        {
            value = null;
            CellType ct = sheet.cellType(line, field.srcSlot);

            if (field.mType.Equals("double"))
            {
                if (ct == CellType.CELLTYPE_BLANK || ct == CellType.CELLTYPE_EMPTY)
                {
                    value = "0";
                    return true;
                }
                if (ct == CellType.CELLTYPE_STRING)
                {
                    string str = sheet.readStr(line, field.srcSlot);
                    if (double.TryParse(str, out double d))
                    {
                        value = d.ToString("G");
                        return true;
                    }
                    return false;
                }

                value = sheet.readNum(line, field.srcSlot).ToString("G");
                return true;
            }

            if (field.mType.Equals("int")
                || field.mType.Equals("centimeter")
                || field.mType.Equals("decimeter")
                || field.mType.Equals("ratio")
                || field.mType.Equals("millimetre"))
            {
                if (ct == CellType.CELLTYPE_BLANK || ct == CellType.CELLTYPE_EMPTY)
                {
                    value = "0";
                    return true;
                }
                if (ct == CellType.CELLTYPE_STRING)
                {
                    string str = sheet.readStr(line, field.srcSlot);
                    if (int.TryParse(str, out int d))
                    {
                        value = d.ToString();
                        return true;
                    }
                    return false;
                }

                value = System.Convert.ToInt32(Math.Round(sheet.readNum(line, field.srcSlot))).ToString();
                return true;
            }

            if (field.mType.Equals("string"))
            {
                if (ct == CellType.CELLTYPE_NUMBER)
                {
                    value = sheet.readNum(line, field.srcSlot).ToString();
                    return true;
                }
                if (ct == CellType.CELLTYPE_STRING)
                {
                    value = sheet.readStr(line, field.srcSlot) ?? string.Empty;
                    return true;
                }
                if (ct == CellType.CELLTYPE_BLANK || ct == CellType.CELLTYPE_EMPTY)
                {
                    value = string.Empty;
                    return true;
                }

                return false;
            }

            string enumKey = sheet.readStr(line, field.srcSlot);
            if (enumKey == null)
            {
                return false;
            }

            value = enumList.GetEnumValue(field.mType, enumKey);
            return value != null;
        }

        private ReadRawLineResult TrySkipDuplicatedPrimary(string filename, SheetCache sheet, int line)
        {
            if (Program.i18nExtraOnly)
            {
                return ReadRawLineResult.Success;
            }

            for (int i = 0; i < fieldConfig.excelFields.Count; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                if (!field.isPrimary || !field.ignore_duplicated)
                {
                    continue;
                }

                if (!TryReadPrimaryKeyForDuplicateCheck(filename, sheet, line, field, out string primary))
                {
                    return ReadRawLineResult.Success;
                }

                if (primarys.Contains(primary))
                {
                    GlobeWarning.Push(string.Format("解析{0}, 行:{1}, 列:{2}, 键值:{3}, 主键={4} 重复，已忽略该行数据",
                        filename, line + 1, field.srcSlot + 1, field.key, primary));

                    if (hints.TryGetValue(primary, out string hit))
                    {
                        GlobeWarning.Push(string.Format("已有项: {0}", hit));
                    }

                    return ReadRawLineResult.Skipped;
                }
            }

            return ReadRawLineResult.Success;
        }

        public ReadRawLineResult ReadExcelRawLine(
            string filename, SheetCache sheet, int line, 
            ref string _key,
            ref List<CellDataForLua> clientLine, 
            ref List<string> serverLine)
        {
            ReadRawLineResult duplicatedPrimaryCheck = TrySkipDuplicatedPrimary(filename, sheet, line);
            if (duplicatedPrimaryCheck == ReadRawLineResult.Skipped)
            {
                return ReadRawLineResult.Skipped;
            }

            for (int i = 0; i < fieldConfig.excelFields.Count; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                CellType ct = sheet.cellType(line, field.srcSlot);

                string s = null;
                uint stringIndex = 0;

                if (Program.i18nExtraOnly)
                {
                    if (!field.mType.Equals("string") || !field.is_text || field.raw_string)
                    {
                        continue;
                    }
                }


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
                            return ReadRawLineResult.Error;
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
                                filename, line+1, field.srcSlot+1, str));
                            return ReadRawLineResult.Error;
                        }
                    }
                    else
                    {
                        d = System.Convert.ToInt32(Math.Round(sheet.readNum(line, field.srcSlot)));
                        s = d.ToString();
                    }
                    CompareNumIsTrue(field, line, filename, d);

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
                            return ReadRawLineResult.Error;
                        }
                    }
                    else
                    {                    
                        d = System.Convert.ToInt32(Math.Round(sheet.readNum(line, field.srcSlot)));
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
                            s = string.Empty;
                        }
                        else
                        {
                            /*
                            //对于可读文本内容，需要进行国际化处理
                            if (field.is_text)
                            {
                                uint textIndex = I18N.RegisterText(s, false);
                                if (textIndex > 0)
                                {
                                    stringIndex = textIndex;
                                }
                            }*/

                            if (!field.raw_string)
                            {
                                uint textIndex = I18N.RegisterText(s, false);
                                if (textIndex > 0)
                                {
                                    stringIndex = textIndex;
                                }
                            }
                        }
                    }
                    else if (ct == CellType.CELLTYPE_BLANK || ct == CellType.CELLTYPE_EMPTY)
                    {
                        s = string.Empty;
                    }
                    else
                    {
                        GlobeError.Push(string.Format("解析{0}错误, 行:{1}, 列:{2}, 单元格数据错误",
                            filename, line+1, field.srcSlot+1));
                        return ReadRawLineResult.Error;
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
                            filename, line+1, field.srcSlot+1, cet));
                        return ReadRawLineResult.Error;
                    }

                    s = enumList.GetEnumValue(field.mType, enumKey);
                    
                    if (s == null)
                    {
                        GlobeError.Push(string.Format("解析{0}错误, 行:{1}, 列:{2}, 键值:{3}, 类型={4}, 枚举={5}, 枚举读取失败",
                            filename, line+1, field.srcSlot+1, field.key, field.mType, enumKey));
                        return ReadRawLineResult.Error;
                    }
                }

                if (s != null)
                {
                    //依赖关系检查
                    if (field.ref_table != null && field.ref_column != null)
                    {
                        if (field.ref_table.Length > 0 && field.ref_column.Length > 0)
                        {
                            if (allLoadCfgData.ContainsKey(field.ref_table) )
                            {
                                allLoadCfgData.TryGetValue(field.ref_table, out Dictionary<string, Dictionary<int, string>> tmpColumnDic);
                                if (tmpColumnDic.ContainsKey(field.ref_column))
                                {
                                    tmpColumnDic.TryGetValue(field.ref_column, out Dictionary<int, string> tmpColumnList);

                                    int num = 0;
                                    foreach (var item in tmpColumnList)
                                    {
                                        if (s == item.Value)
                                        {
                                            num = num + 1;
                                        }
                                    }
                                    if (num == 0)
                                    {
                                        GlobeError.Push(string.Format("当前检测的文件 {0}, 配置项的列名 {1}\n\n " +
                                        "该Excel列名: {2}, 该Excel行号: {3}, 此单元格的值为: {4}, \n\n, 依赖的表名: {5}, 依赖的列名: {6} 该依赖列没有检测到输入值\n",
                                        filename, field.name, field.key, line + 1, s, field.ref_table, field.ref_column));
                                    }
                                }
                            }
                        }
                    }
                }

                if (Program.i18nExtraOnly)
                {
                    continue;
                }

                //主键值（索引Key）
                if (field.isPrimary)
                {
                    if (primarys.Contains(s))
                    {
                        if (field.ignore_duplicated)
                        {
                            GlobeWarning.Push(string.Format("解析{0}, 行:{1}, 列:{2}, 键值:{3}, 主键={4} 重复，已忽略该行数据",
                                filename, line + 1, field.srcSlot + 1, field.key, s));

                            if (hints.TryGetValue(s, out string duplicatedHint))
                            {
                                GlobeWarning.Push(string.Format("已有项: {0}", duplicatedHint));
                            }

                            return ReadRawLineResult.Skipped;
                        }

                        GlobeError.Push(string.Format("解析{0}错误, 行:{1}, 列:{2}, 键值:{3}, 主键={4} 重复!",
                            filename, line+1, field.srcSlot+1, field.key, s));

                        if (hints.TryGetValue(s, out string hit))
                        {
                            GlobeError.Push("已有项: " + hit);
                        }

                        return ReadRawLineResult.Error;
                    }
                    else
                    {
                        _key = s;
                        primarys.Add(s);
                        if (!hints.ContainsKey(s))
                        {
                            hints.Add(s, string.Format("文件:{0}, 行:{1}, 列:{2}", filename, line + 1, field.srcSlot + 1));
                        }
                    }
                }


                //添加到返回结果

                //客户端用的Lua数据
                if (stringIndex > 0)
                {
                    //如果是词条索引类型，需要转换成词条索引的数据结构
                    clientLine.Add(new CellDataForLua(stringIndex, s));
                }
                else
                {
                    clientLine.Add(new CellDataForLua(s));
                }

                //服务器用的CSV数据
                if (!field.client_only)
                {
                    if (null != translatedCsvText && translatedCsvText.TryGetValue(s, out string translated))
                    {
                        if (stringIndex > 0)
                        {
                            //如果是词条索引类型，说明需要翻译
                            serverLine.Add(translated);
                        }
                        else
                        {
                            serverLine.Add(s);
                        }
                    }
                    else
                    {
                        serverLine.Add(s);
                    }
                }
            }

            return ReadRawLineResult.Success;
        }

        private void CompareNumIsTrue(ExcelField field,int line,string filename, int d)
        {
            if (field.target_compare != null && field.self_key != null
                && field.target_key != null && field.ref_table != null)
            {
                if (field.target_compare.Length > 0 && field.self_key.Length > 0
                    && field.target_key.Length > 0 && field.ref_table.Length > 0)
                {
                    string selfCompareKey = GetCompareKey(field.self_key, line, fieldConfig.configName.Substring(0, fieldConfig.configName.IndexOf(".xml")));
                    if (selfCompareKey == null)
                    {
                        GlobeError.Push(string.Format("self_key 无法找到相应的值 请检查当前检测的文件 {0}, 配置项的列名 {1}\n\n " +
                        "该Excel列名: {2}的self_key 是否正确", filename, field.name, field.key));
                        return;
                    }

                    int targetCompareLine = GetTargetCompareLine(field.target_key, field.ref_table, selfCompareKey);
                    if (targetCompareLine == 0)
                    {
                        GlobeError.Push(string.Format("field.target_key 无法找到相应的值 请检查当前检测的文件 {0}, 配置项的列名 {1}\n\n " +
                        "该Excel列名: {2}field.target_key 是否正确", filename, field.name, field.target_key));
                        return;
                    }
                    string compareNum = GetCompareKey(field.target_compare, targetCompareLine, field.ref_table);
                    if (compareNum == null)
                    {
                        GlobeError.Push(string.Format("field.target_compare 无法找到相应的值 请检查当前检测的文件 {0}, 配置项的列名 {1}\n\n " +
                        "该Excel列名: {2}的field.target_compare 是否正确", filename, field.name, field.target_compare));
                        return;
                    }

                    if (field.i_should_be_bigger_than_t)
                    {
                        if (d < int.Parse(compareNum))
                        {
                            GlobeError.Push(string.Format("当前检测的文件 {0}, 配置项的列名 {1}\n\n " +
                            "该Excel列名: {2}, 该Excel行号: {3}, 此单元格的值为: {4}, \n\n, 依赖的表名: {5}, 依赖的列名: {6} 该依赖列不符合规则",
                            filename, field.name, field.key, line + 1, d.ToString(), field.ref_table, field.target_compare));
                        }
                    }
                    else if (field.t_should_be_bigger_than_i)
                    {
                        if (d > int.Parse(compareNum))
                        {
                            GlobeError.Push(string.Format("当前检测的文件 {0}, 配置项的列名 {1}\n\n " +
                            "该Excel列名: {2}, 该Excel行号: {3}, 此单元格的值为: {4}, \n\n, 依赖的表名: {5}, 依赖的列名: {6} 该依赖列不符合规则",
                            filename, field.name, field.key, line + 1, d.ToString(), field.ref_table, field.target_compare));
                        }
                    }
                    else if (field.i_should_like_t)
                    {
                        if (d != int.Parse(compareNum))
                        {
                            GlobeError.Push(string.Format("当前检测的文件 {0}, 配置项的列名 {1}\n\n " +
                            "该Excel列名: {2}, 该Excel行号: {3}, 此单元格的值为: {4}, \n\n, 依赖的表名: {5}, 依赖的列名: {6} 该依赖列不符合规则",
                            filename, field.name, field.key, line + 1, d.ToString(), field.ref_table, field.target_compare));
                        }
                    }
                }
            }
        }

        private string GetCompareKey(string key, int line, string filename)
        {
            if (allLoadCfgData.ContainsKey(filename))
            {
                allLoadCfgData.TryGetValue(filename, out Dictionary<string, Dictionary<int, string>> tmpColumnDic);
                if (tmpColumnDic.ContainsKey(key))
                {
                    tmpColumnDic.TryGetValue(key, out Dictionary<int, string> tmpColumnList);

                    if (tmpColumnList.ContainsKey(line))
                    {
                        string str = tmpColumnList[line];
                        return str;
                    }
                }
            }
            return null;
        }

        private int GetTargetCompareLine( string key, string filename,string keyStr)
        {
            if (allLoadCfgData.ContainsKey(filename))
            {
                allLoadCfgData.TryGetValue(filename, out Dictionary<string, Dictionary<int, string>> tmpColumnDic);
                if (tmpColumnDic.ContainsKey(key))
                {
                    tmpColumnDic.TryGetValue(key, out Dictionary<int, string> tmpColumnList);

                    foreach (var item in tmpColumnList)
                    {
                        if (item.Value == keyStr)
                        {
                            int line = item.Key;
                            return line;
                        }
                    }
                }
            }
            return 0;
        }

    }
}
