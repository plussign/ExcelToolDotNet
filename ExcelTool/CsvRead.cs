using libxl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    public partial class ConvertTool
    {
        public List<string> ReadCsvLine(string filename, SheetCache sheet, int line, ref string _key)
        {
            List<string> content = new List<string>();

            for (int i = 0; i < fieldConfig.excelFields.Count; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                CellType ct = sheet.cellType(line, field.srcSlot);
                string s = null;
                if (field.mType == "double")
                {
                    //if (ct == CellType.CELLTYPE_BLANK || ct == CellType.CELLTYPE_EMPTY)
                    //{
                    //    s = "0";
                    //}
                    //else if (ct == CellType.CELLTYPE_STRING)
                    //{
                    //    string str = sheet.readStr(line, field.srcSlot);
                    //    double d = 0.0f;
                    //    if (double.TryParse(str, out d))
                    //    {
                    //        s = d.ToString("G");
                    //    }
                    //    else
                    //    {
                    //        GlobeError.Push(string.Format(
                    //            "解析{0}单元格错误 行:{1}, 列:{2}, \n约束格式为:浮点数, 输入格式为:字符串, 输入值为: {3}, 无法将该值转换为浮点数",
                    //            filename, line, field.srcSlot, str));
                    //        return null;
                    //    }
                    //}
                    //else
                    //{
                    //    double d = sheet.readNum(line, field.srcSlot);
                    //    s = d.ToString("G");
                    //}
                }
                else if (field.mType == "int")
                {
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
                                filename, line, field.srcSlot, str));
                            return null;
                        }
                    }
                    else
                    {
                        d = (int)sheet.readNum(line, field.srcSlot);
                        s = d.ToString();
                    }
                    CompareNumIsTrue( field, line, filename, d);

                }
                else if (field.mType == "string")
                {
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
                            filename, line, field.srcSlot));
                        return null;
                    }
                }
                else
                {
                    string enumType = sheet.readStr(line, field.srcSlot);
                    if (enumType == null)
                    {
                        GlobeError.Push(string.Format("解析excel文件{0}错误, 无法以字符串方式读取单元格, 行:{1}, 列:{2}",
                            filename, line, field.srcSlot));
                        return null;
                    }

                    s = enumList.GetEnumValue(field.mType, enumType);
                    if (s == null)
                    {
                        GlobeError.Push(string.Format("解析{0}错误, 行:{1}, 列:{2}, 键值:{3}, 类型={4}, 枚举={5}, 枚举读取失败",
                            filename, line, field.srcSlot, field.key, field.mType, enumType));
                        return null;
                    }
                }
                if (s != "")
                {
                    if (field.ref_table != null && field.ref_column != null)
                    {
                        if (field.ref_table.Length > 0 && field.ref_column.Length > 0)
                        {
                            if (allLoadCfgData.ContainsKey(field.ref_table))
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

                if (field.isPrimary)
                {
                    if (primarys.Contains(s))
                    {
                        GlobeError.Push(string.Format("解析{0}错误, 行:{1}, 列:{2}, 键值:{3}, 主键={4} 重复!",
                            filename, line, field.srcSlot, field.key, s));

                        if (hints.TryGetValue(s, out string hit))
                        {
                            GlobeError.Push("已有项: " + hit);
                        }

                        return null;
                    }
                    else
                    {
                        _key = s;
                        primarys.Add(s);
                    }
                }
                content.Add(s);

            }

            return content;
        }

        public List<string> ReadErlCsvLine(string filename, SheetCache sheet, int line)
        {
            List<string> content = new List<string>();
            for (int i = 0; i < fieldConfig.excelFields.Count; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                CellType ct = sheet.cellType(line, field.srcSlot);
                string s = null;
                if (field.client_only)
                {
                    continue;
                }
                else
                {
                    if (field.mType == "double")
                    {
                        if (ct == CellType.CELLTYPE_BLANK || ct == CellType.CELLTYPE_EMPTY)
                        {
                            s = "0";
                        }
                        else if (ct == CellType.CELLTYPE_STRING)
                        {
                            string str = sheet.readStr(line, field.srcSlot);
                            if (double.TryParse(str, out double d))
                            {
                                s = d.ToString("G");
                            }
                            else
                            {
                                return null;
                            }
                        }
                        else
                        {
                            double d = sheet.readNum(line, field.srcSlot);
                            s = d.ToString("G");
                        }
                    }
                    else if (field.mType == "int")
                    {
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
                                return null;
                            }
                        }
                        else
                        {
                            d = (int)sheet.readNum(line, field.srcSlot);
                            s = d.ToString();
                        }

                    }
                    else if (field.mType == "string")
                    {
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
                            return null;
                        }
                    }
                    else
                    {
                        string enumType = sheet.readStr(line, field.srcSlot);
                        if (enumType == null)
                        {
                            return null;
                        }

                        s = enumList.GetEnumValue(field.mType, enumType);
                        if (s == null)
                        {
                            return null;
                        }
                    }
                    content.Add(s);
                }
            }
            return content;
        }
    }
}
