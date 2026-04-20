using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTool
{
    public partial class ConvertTool
    {
        //static Random r = new Random(System.Environment.TickCount);
        private string FormatLuaLineString(List<CellDataForLua> input)
        {
            StringBuilder content = new StringBuilder();
            string key = string.Empty;

            uint[] skippedFields = null;
            int fieldCount = fieldConfig.excelFields.Count;
            if (fieldCount > 2)
            {
                int skippedFieldsRecords = (int)Math.Ceiling((double)fieldCount / 32);
                skippedFields = new uint[skippedFieldsRecords];
                for (int i = 0; i < skippedFieldsRecords; ++i)
                {
                    skippedFields[i] = 0;
                }
            }

            for (int i = 0; i < fieldCount; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                CellDataForLua cellData = input[i];

                bool isStringField = field.mType.Equals("string");
                bool isSkippableKey = (field.isPrimary && !isStringField);

                if (fieldCount <= 2 
                    || !(cellData.IsBlank || isSkippableKey))
                {
                    //不跳过当前单元格
                    string cellString = string.Empty;

                    if (isStringField)
                    {
                        if (CellDataForLua.CellTypeForLua.Standard == cellData.type)
                        {
                            uint textIndex = I18N.RegisterText(cellData.GetOriginalString(), false);
                            if (textIndex > 0)
                            {
                                cellString = string.Format("{0}", textIndex);
                            }
                            else
                            {
                                //未放入词条表的表格内字符串，需要进行Lua字符串格式调整
                                cellString = Assist.ToLuaStr(cellData.GetOriginalString());
                            }
                        }
                        else
                        {
                            //索引词条表
                            uint strId = cellData.GetStringIndex();
                            cellString = string.Format("{0}", strId);
                        }
                    }
                    else
                    {
                        cellString = cellData.GetOriginalString();
                    }

                    if (content.Length > 0)
                    {
                        content.Append(",");
                    }
                    content.Append(cellString);
                }
                else
                {
                    //跳过处理
                    int skippedFieldsRecordIndex = i / 32;
                    skippedFields[skippedFieldsRecordIndex] |= ((uint)0x1 << (i - 32 * skippedFieldsRecordIndex));
                }

                if (field.isPrimary)
                {
                    if (!field.mType.Equals("string"))
                    {
                        key = cellData.GetOriginalString();
                    }
                    else
                    {
                        key = Assist.ToLuaStr(cellData.GetOriginalString());
                    }
                }
            }

            StringBuilder sbSkippedRecords = new StringBuilder();
            if (null != skippedFields)
            {
                for (int i = 0; i < skippedFields.Length; ++i)
                {
                    sbSkippedRecords.Append(skippedFields[i].ToString());
                    sbSkippedRecords.Append(",");
                }
            }

            return string.Format("[{0}]={{{1}{2}}},\r\n", key, sbSkippedRecords.ToString(), content.ToString());
            //return string.Format("[{0}]=\"{1}\",\r\n", key, r.Next(1000, 9999));
        }

        private string FormatLuaLineStringDynamic(List<CellDataForLua> input)
        {
            StringBuilder content = new StringBuilder();
            string key = string.Empty;

            uint[] skippedFields = null;
            int fieldCount = fieldConfig.excelFields.Count;
            if (fieldCount > 2)
            {
                int skippedFieldsRecords = (int)Math.Ceiling((double)fieldCount / 32);
                skippedFields = new uint[skippedFieldsRecords];
                for (int i = 0; i < skippedFieldsRecords; ++i)
                {
                    skippedFields[i] = 0;
                }
            }

            for (int i = 0; i < fieldCount; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                CellDataForLua cellData = input[i];

                bool isStringField = field.mType.Equals("string");

                //不跳过当前单元格
                string cellString = string.Empty;

                if (isStringField)
                {
                    if (CellDataForLua.CellTypeForLua.Standard == cellData.type)
                    {
                        uint textIndex = I18N.RegisterText(cellData.GetOriginalString(), false);
                        if (textIndex > 0)
                        {
                            cellString = Assist.ToLuaStr("$$" + textIndex);
                        }
                        else
                        {
                            //未放入词条表的表格内字符串，需要进行Lua字符串格式调整
                            cellString = Assist.ToLuaStr(cellData.GetOriginalString());
                        }
                    }
                    else
                    {
                        //索引词条表
                        uint strId = cellData.GetStringIndex();
                        cellString = Assist.ToLuaStr("$$" + strId);
                    }
                }
                else
                {
                    cellString = cellData.GetOriginalString();
                }


                cellString = string.Format("{0} = {1}", field.name.ToLower(), cellString);
                if (content.Length > 0)
                {
                    content.Append(",");
                }
                content.Append(cellString);

                if (field.isPrimary)
                {
                    if (!field.mType.Equals("string"))
                    {
                        key = cellData.GetOriginalString();
                    }
                    else
                    {
                        key = Assist.ToLuaStr(cellData.GetOriginalString());
                    }
                }
            }

            return string.Format("[{0}]={{{1}}},\r\n", key, content.ToString());
        }
    }
}
