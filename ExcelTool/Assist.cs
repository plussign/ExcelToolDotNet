using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    class Assist
    {
        static public string GetCppRtString(string mType)
        {
            if (mType == "string")
            {
                return "fs.ReadString();";
            }
            else if (mType == "number" || mType == "double")
            {
                return "fs.ReadFloat();";
            }

            return "fs.ReadInt();";
        }

        static public string GetCSRtString(string mType)
        {
            if (mType == "string")
            {
                return "br.ReadString();";
            }
            else if (mType == "number" || mType == "double")
            {
                return "br.ReadSingle();";
            }
            else if (mType == "centimeter")
            {
                return "Fix.Ratio( br.ReadInt32(),100);";
            }
            else if (mType == "decimeter")
            {
                return "Fix.Ratio( br.ReadInt32(),10);";
            }
            else if (mType == "millimetre")
            {
                return "Fix.Ratio( br.ReadInt32(),1000);";
            }
            else if (mType == "ratio")
            {
                return "Fix.Ratio( br.ReadInt32(),10000);";
            }

            return "br.ReadInt32();";
        }

        static public string GetCppTypeByStr(string mType)
        {
            if (mType == "string")
            {
                return "std::string";
            }
            else if (mType == "number" || mType == "double")
            {
                return "double";
            }

            return "int";
        }

        static public string GetCSharpTypeByStr(string mType)
        {
            if (mType == "string")
            {
                return "string";
            }
            else if (mType == "number" || mType == "double")
            {
                return "double";
            }
            else if (mType == "centimeter" || mType == "decimeter" || mType == "ratio" || mType == "millimetre")
            {
                return "Fix";
            }

            return "int";
        }

        static public string ToCppStr(string str)
        {
            str = str.Replace("\\", "\\\\");
            str = str.Replace("\"", "\\\"");
            str = str.Replace("\n", "");
            return string.Format("\"{0}\"", str);
        }

        static public string ToLuaStr(string str)
        {
            while (true)
            {
                int pos = str.IndexOf('\\');
                if (pos >= 0)
                {
                    break;
                }

                pos = str.IndexOf("'");
                if (pos >= 0)
                {
                    break;
                }

                pos = str.IndexOf("\"");
                if (pos >= 0)
                {
                    break;
                }

                pos = str.IndexOf("\n");
                if (pos >= 0)
                {
                    break;
                }

                return string.Format("\"{0}\"", str);
            }

            // 首尾添加空格, 防止使用该值做为 Key 
            // [Key]=[[[str]]] 非法
            // [ Key ]=[ [[str]] ] 合法
            return string.Format(" [=[{0}]=]", str);
        }

        static public string ToGoStr(string str)
        {
            /*
            while (true)
            {
                int pos = str.IndexOf('\\');
                if (pos >= 0)
                {
                    break;
                }

                pos = str.IndexOf("'");
                if (pos >= 0)
                {
                    break;
                }

                pos = str.IndexOf("\"");
                if (pos >= 0)
                {
                    break;
                }

                pos = str.IndexOf("\n");
                if (pos >= 0)
                {
                    break;
                }

                return string.Format("`{0}`", str);
            }*/
            
            return string.Format("`{0}`", str);
        }
    }
}
