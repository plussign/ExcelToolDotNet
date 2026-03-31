using System.Collections.Generic;
using System.Xml;
using System.IO;

namespace ExcelTool
{
    public class ExcelField
    {
        public string key;
        public string mType;
        public int srcSlot = 0;
        public bool isPrimary = false;
        public bool enum_value = false;
        public string name;
        public int index = 0;
        public bool needLoadIntoMemory = false;
        public string ref_table;
        public string ref_column;
        public bool client_only = false;
        public bool export_bin = false;
        public bool is_text = false;    //是否是可读文本，如果是需要做国际化处理
        public bool sdf_text = false;
        public string min_num;
        public string max_num;
        public bool i_should_be_bigger_than_t = false;
        public bool t_should_be_bigger_than_i = false;
        public bool i_should_like_t = false;
        public string self_key;
        public string target_key;
        public string target_compare;
        public string bin_to;

        public bool skip_export_bin
        {
            get 
            {
                //if (Program.outputCSharpAccessInterface)
                //{
                //    return false;
                //}

                return (!isPrimary && !export_bin); 
            }
        }

        public bool is_bin_lua
        {
            get
            {
                var low_bin = bin_to.ToLower();
                return low_bin == "all" || low_bin == "lua";
            }
        }

        public bool is_bin_cs
        {
            get
            {
                var low_bin = bin_to.ToLower();
                return low_bin == "all" || low_bin == "csharp";
            }
        }
    }

    public class InputConfig
    {
        public struct SourceFileInfo
        {
            public string fileName;
            public bool sourceDynamic;
        }
        public List<SourceFileInfo> files = new List<SourceFileInfo>();
        
        // 特殊大区
        private string special_channel = string.Empty;

        public InputConfig(string special)
        {
            special_channel = special;
        }

        private void AddFile(string fname, bool sourceDynamic = false)
        {
            if (!string.IsNullOrEmpty(special_channel))
            {
                string pathname = string.Format("{0}/source/{1}", special_channel, fname);
                if (File.Exists(pathname))
                {
                    SourceFileInfo info = new SourceFileInfo
                    {
                        fileName = pathname,
                        sourceDynamic = sourceDynamic
                    };
                    files.Add(info);
                    Log.WriteLine("添加特别表格: {0}", pathname);
                    return;
                }
            }

            string defFilename = string.Format("source/{0}", fname);
            SourceFileInfo defInfo = new SourceFileInfo
            {
                fileName = defFilename,
                sourceDynamic = sourceDynamic
            };
            files.Add(defInfo);
        }

        public bool Load(XmlElement root)
        {
            foreach (XmlElement node in root)
            {
                if (node.Name == "input")
                {
                    foreach (XmlElement child in node)
                    {
                        string fname = child.GetAttribute("file");
                        fname = fname.Replace("\\", "/");
                        bool sourceDynamic = (child.GetAttribute("dynamic") == "true");
                        AddFile(fname, sourceDynamic);
                    }

                    return true;
                }
            }

            return false;
        }
    }

    public class OutputConfig
    {
        public string filename;

        public bool Load(XmlElement root)
        {
            foreach (XmlElement node in root)
            {
                if (node.Name == "output")
                {
                    filename = node.GetAttribute("file");
                    return true;
                }
            }

            return false;
        }
    }

    public class EnumItem
    {
        public string text;
        public string luaName;
        public string value;
    }
}
