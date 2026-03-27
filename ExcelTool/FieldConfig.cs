using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace ExcelTool
{
    public enum TableUseMode
    {
        Common,
        Client,
        Server,
    }

    public partial class FieldConfig
    {
        public List<ExcelField> excelFields = new List<ExcelField>();
        public string tableName;
        public string tableDesc;
        public string configName;
        public string export_csharp;
        public string export_golang;
        public string export_xml;
        public string export_erl;
        public bool export_enum_only;
        public TableUseMode use_mode = TableUseMode.Common;

        public string GetCppPrimaryKey()
        {
            for(int i=0; i<excelFields.Count; ++i)
            {
                if (excelFields[i].isPrimary)
                {
                    return Assist.GetCppTypeByStr(excelFields[i].mType);
                }
            }
            return "";
        }

        public string GetCSharpPrimaryKey()
        {
            for (int i = 0; i < excelFields.Count; ++i)
            {
                if (excelFields[i].isPrimary)
                {
                    return Assist.GetCSharpTypeByStr(excelFields[i].mType);
                }
            }
            return "";
        }

        public bool LoadFields(XmlElement root)
        {
            int index = 0;
            XmlNode rootNode = root.FirstChild;
            while (rootNode != null)
            {
                ExcelField field = new ExcelField();
                if (rootNode is XmlElement node)
                {
                    field.key = node.GetAttribute("key");

                    field.mType = node.GetAttribute("type");
                    field.isPrimary = node.GetAttribute("primary") == "true";
                    field.name = node.GetAttribute("name");
                    field.enum_value = node.GetAttribute("enum_value") == "true";
                    field.needLoadIntoMemory = node.GetAttribute("need_load") == "true";
                    field.ref_table = node.GetAttribute("ref_table");
                    field.ref_column = node.GetAttribute("ref_column");
                    field.client_only = node.GetAttribute("client_only") == "true";
                    field.export_bin = node.GetAttribute("export_bin") == "true";
                    field.is_text = node.GetAttribute("text") == "true";
                    field.sdf_text = node.GetAttribute("sdf_text") == "true";
                    field.min_num = node.GetAttribute("min_num");
                    field.max_num = node.GetAttribute("max_num");
                    field.i_should_be_bigger_than_t = node.GetAttribute("i_should_be_bigger_than_t") == "true";
                    field.t_should_be_bigger_than_i = node.GetAttribute("t_should_be_bigger_than_i") == "true";
                    field.i_should_like_t = node.GetAttribute("i_should_like_t") == "true";
                    field.self_key = node.GetAttribute("self_key");
                    field.target_key = node.GetAttribute("target_key");
                    field.target_compare = node.GetAttribute("target_compare");
                    //field.bin_to = node.GetAttribute("bin_to");

                    field.index = index++;

                    if (field.key.Length == 0)
                    {
                        GlobeError.Push(string.Format("key没有定义!"));
                        return false;
                    }

                    /*
                    if (field.bin_to.Length == 0)
                    {
                        GlobeError.Push(string.Format("bin_to没有定义!"));
                       // return false;
                    }*/

                    if (field.mType.Length == 0)
                    {
                        GlobeError.Push(string.Format("type没有定义, key={0}", field.key));
                        return false;
                    }

                    if (field.name.Length == 0)
                    {
                        GlobeError.Push(string.Format("name没有定义, key={0}", field.key));
                        return false;
                    }

                    excelFields.Add(field);
                }

                rootNode = rootNode.NextSibling;
                              
            }
            return true;
        }

        public bool hasHumanReadableText()
        {
            foreach (ExcelField field in excelFields)
            {
                if (field.is_text)
                {
                    return true;
                }
            }

            return false;
        }

        public string GetPramaryKeyName()
        {
            for (int i = 0; i < excelFields.Count; ++i)
            {
                var field = excelFields[i];
                if (field.isPrimary)
                {
                    return field.name.ToLower();
                }
            }
            return string.Empty;
        }

        public string TypeToErl(string mType)
        {
            if (mType.Equals("string"))
            {
                return "string()";
            }
            else if (mType.Equals("double"))
            {
                return "float()";
            }
            else
            {
                return "integer()";
            }
        }

        public string TypeToGo(string mType)
        {
            if (mType.Equals("string"))
            {
                return "string";
            }
            else if (mType.Equals("double"))
            {
                return "float64";
            }
            else
            {
                return "int";
            }
        }

        public string TypeToFun(string mType)
        {
            if (mType.Equals("string"))
            {
                return "to_list";
            }
            else if (mType.Equals("double"))
            {
                return "to_float";
            }
            else
            {
                return "to_int";
            }
        }

        public bool Load(XmlElement root, string configXMLFileName)
        {
            tableName = root.GetAttribute("name");
            tableDesc = root.GetAttribute("desc");
            export_csharp = root.GetAttribute("export_csharp");
            export_xml = root.GetAttribute("export_xml");
            export_golang = root.GetAttribute("export_golang");
            string s_export_enum_only = root.GetAttribute("export_enum_only");
            export_enum_only = !(s_export_enum_only.Length == 0 || s_export_enum_only == "0");
            export_erl = root.GetAttribute("export_erl");
            string str = root.GetAttribute("use_mode").ToLower();
            if (str == "client")
            {
                use_mode = TableUseMode.Client;
            }
            else if (str == "server")
            {
                use_mode = TableUseMode.Server;
            }
            else
            {
                use_mode = TableUseMode.Common;
            }

            configName = configXMLFileName;

            if (tableName.Length == 0)
            {
                GlobeError.Push(string.Format("table.name没有定义"));
                return false;
            }

            foreach (XmlElement node in root)
            {
                if (node.Name == "fields")
                {
                    return LoadFields(node);
                }
            }

            return false;
        }

        public int GetSlot(SheetCache sheet, string text, string filename)
        {
            int num = sheet.lastCol();
            for (int i = 0; i < num; ++i)
            {
                string str = sheet.readStr(0, i);
                if (str != null && str == text)
                {
                    return i;
                }
            }

            GlobeError.Push(string.Format("当前处理 \"{0}\"\n在Excel表 \"{1}\" 中查找列 \"{2}\" 失败!", configName, filename, text));

            return -1;
        }

        public bool LoadSlotInfo(SheetCache sheet, string filename)
        {
            int num = sheet.lastCol();
            foreach (ExcelField field in excelFields)
            {
                int slot = GetSlot(sheet, field.key, filename);
                if (slot == -1)
                {
                    return false;
                }
                field.srcSlot = slot;
            }

            return true;
        }
    }
}
