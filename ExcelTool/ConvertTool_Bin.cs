using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTool
{
    public partial class ConvertTool
    {
        public string FormatBinReadText(string filename)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(string.Format("\tZBDataTableReader fs(\"{0}\");\r\n\tif (!fs.isValid()) return false;\r\n\tint32 head=fs.ReadInt();\r\n", filename));
            sb.Append("\tif (head!=12345678) return false;\r\n\tint32 __count=fs.ReadInt();\r\n\tfor(int32 i=0; i<__count; ++i)\r\n\t{\r\n");

            StringBuilder args = new StringBuilder();
            args.Append(string.Format("\t\tdata_{1}[{0}] = new Table_{1} (", GetPramaryKeyName(), fieldConfig.tableName));

            for (int i = 0; i < fieldConfig.excelFields.Count; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                if (field.skip_export_bin)
                {
                    continue;
                }

                sb.Append(string.Format("\t\t{0} {1} = {2}\r\n", Assist.GetCppTypeByStr(field.mType), field.name.ToLower(), Assist.GetCppRtString(field.mType)));

                if (i != 0)
                {
                    args.Append(",");
                }
                args.Append(field.name.ToLower());
            }

            args.Append(");\r\n");
            sb.Append(args);
            sb.Append("\t}\r\n\treturn true;");
            return sb.ToString();
        }

        public byte[] FormatBinLine(List<CellDataForLua> input, string currType)
        {
            List<byte[]> allLine = new List<byte[]>();

            for (int i = 0; i < fieldConfig.excelFields.Count; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                if (field.skip_export_bin)
                {
                    continue;
                }

                if (field.mType == "string")
                {
                    byte[] bytes = System.Text.Encoding.UTF8.GetBytes(input[i].GetOriginalString());
                    int len = bytes.Length;
                    byte[] buf2 = BitConverter.GetBytes(len);
                    byte b1 = (byte)(len >> 24);

                    allLine.Add(buf2);
                    allLine.Add(bytes);

                }
                else if (field.mType == "double" || field.mType == "number")
                {
                    byte[] buf = BitConverter.GetBytes(float.Parse(input[i].GetOriginalString()));
                    allLine.Add(buf);
                }
                else
                {
                    byte[] buf = BitConverter.GetBytes(int.Parse(input[i].GetOriginalString()));
                    allLine.Add(buf);
                }
            }

            return BaseHelper.Meger(allLine);
        }
    }
}
