using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace ExcelTool
{
    public class EnumManager
    {
        public Dictionary<string, Dictionary<string, EnumItem>> items
            = new Dictionary<string, Dictionary<string, EnumItem>>();

        public string ExportCSharpCode()
        {
            StringBuilder allEnumText = new StringBuilder();

            allEnumText.Append(@"
public static class TableEnum
{ 
");

            foreach (var _v1 in items)
            {
                allEnumText.AppendFormat(@"
    public enum {0}
    {{", _v1.Key);

                foreach (var _v2 in _v1.Value)
                {
                    allEnumText.AppendFormat(@"
        {0} = {1}, //{2}", _v2.Value.luaName, _v2.Value.value, _v2.Value.text);
                }

                allEnumText.Append(@"
    };
");
            }

            allEnumText.Append(@"
    public enum CommonEnums
    {");
            foreach (var kv in CustomEnumMgr.Enums)
            {
                allEnumText.AppendFormat(@"
        {0}={1},", kv.Key, kv.Value);
            }
            allEnumText.Append(@"
    };
};");

            return allEnumText.ToString();
        }

        public string ExportCPPCode()
        {
            StringBuilder allEnumText = new StringBuilder();

            allEnumText.Append("#pragma once\r\n");

            foreach (var _v1 in items)
            {
                string k = string.Format("\r\nenum {0}\r\n{{\r\n", _v1.Key);
                allEnumText.Append(k);

                foreach (var _v2 in _v1.Value)
                {
                    string str = string.Format("    {0}={1}, //{2}\r\n",
                        _v2.Value.luaName, _v2.Value.value, _v2.Value.text);
                    allEnumText.Append(str);
                }

                allEnumText.Append("};\r\n");
            }

            allEnumText.Append("\r\nenum CommonEnums\r\n{\r\n");
            allEnumText.Append(CustomEnumMgr.ToCpp());
            allEnumText.Append("};\r\n");

            return allEnumText.ToString();
        }

        public string ExportLuaCode()
        {
            StringBuilder allEnumText = new StringBuilder();
            foreach (var _v1 in items)
            {
                string newLine = string.Format("\n--------------------\n--- {0} \n--------------------\n", _v1.Key);
                allEnumText.Append(newLine);

                foreach (var _v2 in _v1.Value)
                {
                    string str = string.Format("{0}={1} ---{2}\n",
                        _v2.Value.luaName, _v2.Value.value, _v2.Value.text);
                    allEnumText.Append(str);
                }
            }

            allEnumText.Append("\n------Common enums------\n");
            allEnumText.Append(CustomEnumMgr.ToLua());

            return allEnumText.ToString();
        }

        public string ExportGoCode()
        {
            StringBuilder allEnumText = new StringBuilder();
            allEnumText.Append("package csv\n\n");
            foreach (var _v1 in items)
            {
                string newLine = string.Format("const (\n");
                allEnumText.Append(newLine);

                foreach (var _v2 in _v1.Value)
                {
                    string str = string.Format("\t{0} = {1}\n",
                        _v2.Value.luaName, _v2.Value.value);
                    allEnumText.Append(str);
                }
                allEnumText.Append(")\n");
            }

            return allEnumText.ToString();
        }

        

        public string ExportErlCode()
        {
            string allEnumText = "";
            foreach (var _v1 in items)
            {
                foreach (var _v2 in _v1.Value)
                {
                    string str = string.Format("-define({0}, {1}).\n", _v2.Value.luaName, _v2.Value.value);
                    allEnumText += str;
                }
            }

            return allEnumText;
        }

        private void LoadEnumNode(XmlLinkedNode rootNode)
        {
            Dictionary<string, EnumItem> kv = new Dictionary<string, EnumItem>();

            foreach (XmlLinkedNode linkedNode in rootNode)
            {
                XmlElement node = linkedNode as XmlElement;
                if (null != node)
                {
                    string key = node.GetAttribute("key");
                    string value = node.GetAttribute("value");
                    string name = node.GetAttribute("name");

                    EnumItem item = new EnumItem
                    {
                        text = key,
                        value = value,
                        luaName = name
                    };
                    if (!kv.ContainsKey(key))
                    {
                        kv.Add(key, item);
                    }
                    else
                    {
                        GlobeError.Push(string.Format("添加枚举:[{0}],重复!", key));
                    }
                }
            }

            XmlElement root = rootNode as XmlElement;
            string enum_name = root.GetAttribute("name");

            if (!items.ContainsKey(enum_name))
            {
                items.Add(enum_name, kv);
            }
            else
            {
                GlobeError.Push(string.Format("添加枚举项:[{0}],重复!", enum_name));
            }
        }

        private void LoadEnumType(XmlElement root)
        {
            if (null != root)
            {
                foreach (XmlElement node in root)
                {
                    if (null != node)
                    {
                        LoadEnumNode(node);
                    }
                }
            }
        }

        public string GetEnumValue(string typeName, string enumName)
        {
            if (items.TryGetValue(typeName, out Dictionary<string, EnumItem> o))
            {
                if (o.TryGetValue(enumName, out EnumItem s))
                {
                    return s.value;
                }
            }
            return null;
        }

        public string GetEnumLuaName(string typeName, string enumName)
        {
            if (items.TryGetValue(typeName, out Dictionary<string, EnumItem> o))
            {
                if (o.TryGetValue(enumName, out EnumItem s))
                {
                    return s.luaName;
                }
            }
            return null;
        }

        public bool Load(XmlElement root)
        {
            foreach (XmlElement node in root)
            {
                if (node.Name == "enums")
                {
                    LoadEnumType(node);
                    return true;
                }
            }

            return false;
        }
    }

}
