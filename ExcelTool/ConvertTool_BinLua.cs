using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelTool
{
    public partial class ConvertTool
    {
        private byte[] FormatBinLuaLineString(List<CellDataForLua> input, bool iskeyStr)
        {
            string key = string.Empty;
            int keyint = 0;

            int fieldCount = fieldConfig.excelFields.Count;

            bool useFieldsSkip = (fieldCount > 2);
            uint[] skippedFields = null;

            if (useFieldsSkip)
            {
                int skippedFieldsRecords = (int)Math.Ceiling((double)fieldCount / 32);
                skippedFields = new uint[skippedFieldsRecords];
                for (int i = 0; i < skippedFieldsRecords; ++i)
                {
                    skippedFields[i] = 0;
                }
            }
            //             else
            //             {
            //                 Console.Out.WriteLine(fieldConfig.tableName);
            //             }

            MemoryStream memContent = new MemoryStream(300);
            int fieldNum = 0;
            for (int i = 0; i < fieldCount; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                CellDataForLua cellData = input[i];

                bool isStringField = field.mType.Equals("string");
                bool isSkippableKey = (field.isPrimary && !isStringField);
                
                //if (field.is_bin_lua )
                {
                    if (!useFieldsSkip || !(cellData.IsBlank || isSkippableKey))
                    {
                        //不跳过当前单元格
                        string cellString = string.Empty;

                        if (isStringField)
                        {
                            if (field.raw_string)
                            {
                                //raw_string字段直接写入UTF8字符串
                                byte[] strBytes = Encoding.UTF8.GetBytes(cellData.GetOriginalString());
                                memContent.Write(BitConverter.GetBytes(strBytes.Length), 0, 4);
                                memContent.Write(strBytes, 0, strBytes.Length);
                            }
                            else
                            {
                                //普通字符串字段写入字符串表索引
                                uint textIndex;
                                if (CellDataForLua.CellTypeForLua.Standard == cellData.type)
                                {
                                    textIndex = I18N.RegisterText(cellData.GetOriginalString(), false);
                                }
                                else
                                {
                                    textIndex = cellData.GetStringIndex();
                                }
                                memContent.Write(BitConverter.GetBytes((int)textIndex), 0, 4);
                            }
                        }
                        else if (field.mType.Equals("double"))
                        {
                            //double类型
                            string valueStr = cellData.IsBlank ? "0" : cellData.GetOriginalString();
                            if (float.TryParse(valueStr, out float floatValue))
                            {
                                memContent.Write(BitConverter.GetBytes(floatValue), 0, 4);
                            }
                            else
                            {
                                throw new Exception("无法转换float值");
                            }
                        }
                        else
                        {
                            //int double 或者枚举值类型的单元格
                            string valueStr = cellData.IsBlank ? "0" : cellData.GetOriginalString();
                            if (int.TryParse(valueStr, out int intValue))
                            {
                                if (field.mType.Equals("centimeter"))
                                {
                                    memContent.Write(BitConverter.GetBytes((float)intValue/100), 0, 4);
                                }
                                else if (field.mType.Equals("decimeter"))
                                {
                                    memContent.Write(BitConverter.GetBytes((float)intValue / 10), 0, 4);
                                }
                                else if (field.mType.Equals("millimetre"))
                                {
                                    memContent.Write(BitConverter.GetBytes((float)intValue / 1000), 0, 4);
                                }
                                else if (field.mType.Equals("ratio"))
                                {
                                    memContent.Write(BitConverter.GetBytes((float)intValue / 10000), 0, 4);
                                }
                                else
                                {
                                    memContent.Write(BitConverter.GetBytes(intValue), 0, 4);
                                }
                            }
                            else
                            {
                                throw new Exception("无法转换int值");
                            }
                        }
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
                            if (!int.TryParse(cellData.GetOriginalString(), out keyint))
                            {
                                throw new Exception("无法转换int值");
                            }
                        }
                        else
                        {
                            key = cellData.GetOriginalString();
                            if ("SESSION_PLAYER_PREPARE" == key)
                            {
                                string kfoa = key;
                            }
                        }
                    }

                    fieldNum = fieldNum + 1;
                }

            }

            MemoryStream memSkippedRecords = new MemoryStream(20);
            if (useFieldsSkip)
            {
                for (int i = 0; i < skippedFields.Length; ++i)
                {
                    memSkippedRecords.Write(BitConverter.GetBytes(skippedFields[i]), 0, 4);
                }
            }

            MemoryStream allLine = new MemoryStream(500);

            if (fieldNum <= 1)
            {
                return null;
            }

            UTF8Encoding encoding = new UTF8Encoding();

            if(iskeyStr)
            {
                byte[] keyBytes = encoding.GetBytes(key);
                allLine.Write(BitConverter.GetBytes((ushort)keyBytes.Length), 0, 2);
                allLine.Write(keyBytes, 0, keyBytes.Length);
            }
            else
            {
                allLine.Write(BitConverter.GetBytes(keyint), 0, 4);
            }

            ushort len = (ushort)memContent.Length;
            if(useFieldsSkip)
            {
                len += (ushort)memSkippedRecords.Length;
            }
            allLine.Write(BitConverter.GetBytes(len), 0, 2);
            if (useFieldsSkip)
            {
                allLine.Write(memSkippedRecords.GetBuffer(), 0, (int)memSkippedRecords.Length);
            }

            allLine.Write(memContent.GetBuffer(), 0, (int)memContent.Length);

            allLine.Seek(0, SeekOrigin.Begin);
            byte[] output = new byte[allLine.Length];
            allLine.Read(output, 0, (int)allLine.Length);

            return output;
        }

        private byte[] _buffer = new byte[1024];
        private byte[] FormatBinCSharpLineString(List<CellDataForLua> input)
        {
            int fieldCount = fieldConfig.excelFields.Count;
            int offset = 0;
            using (BinaryWriter bw = new BinaryWriter(new MemoryStream(_buffer)))
            {
                //int fieldNum = 0;
                for (int i = 0; i < fieldCount; ++i)
                {
                    ExcelField field = fieldConfig.excelFields[i];

                    if (field.skip_export_bin)
                        continue;

                    //if (!field.is_bin_cs && !field.isPrimary)
                    //    continue;

                    CellDataForLua cellData = input[i];

                    bool isStringField = field.mType.Equals("string");

                    if (isStringField)
                    {
                        bw.Write(cellData.GetOriginalString());
                    }
                    else
                    {
                        //int double 或者枚举值类型的单元格
                        if (!cellData.IsBlank)
                        {
                            if (!field.mType.Equals("double"))
                            {

                                if (int.TryParse(cellData.GetOriginalString(), out int intValue))
                                {
                                    bw.Write(intValue);
                                }
                                else
                                {
                                    throw new Exception("无法转换int值");
                                }
                            }
                            else
                            {
                                if (float.TryParse(cellData.GetOriginalString(), out float floatValue))
                                {
                                    bw.Write(floatValue);
                                }
                                else
                                {
                                    throw new Exception("无法转换float值");
                                }
                            }
                        }
                        else
                        {
                            if (field.mType.Equals("double"))
                                bw.Write(0f);
                            else
                                bw.Write(0);
                        }
                    }
                }
                offset = (int)bw.BaseStream.Position;
            }

            byte[] output = new byte[offset];
            Buffer.BlockCopy(_buffer, 0, output, 0, offset);

            return output;
        }

        private void SaveBinLuaList(List<string> allOutName ,string filePath)
        {
            MemoryStream allLine = new MemoryStream(500 * 40);
            allLine.Write(BitConverter.GetBytes((ushort)allOutName.Count), 0, 2);
            foreach (var name in allOutName)
            {
                var data = UTF8Encoding.Default.GetBytes(name);
                allLine.Write(BitConverter.GetBytes((ushort)data.Length), 0, 2);
                allLine.Write(data, 0, data.Length);

            }
            allLine.Seek(0, SeekOrigin.Begin);
            byte[] output = new byte[allLine.Length];
            allLine.Read(output, 0, (int)allLine.Length);
            BaseHelper.WriteBin(filePath, output);
        }
    }
}
