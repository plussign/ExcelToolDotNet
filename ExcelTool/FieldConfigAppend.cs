using System.IO;
using System.Text;

namespace ExcelTool
{
    public partial class FieldConfig
    {

        public void AppendCppDefine(ref string str, string keyType, string outputFilename)
        {
            string head = string.Format("struct Table_{0} {{\n", tableName);

            string construct = string.Format("\tTable_{0}(", tableName);
            for (int i = 0; i < excelFields.Count; ++i)
            {
                ExcelField field = excelFields[i];
                if (field.skip_export_bin)
                {
                    continue;
                }
                string fname = field.name.ToLower();
                string fixType = Assist.GetCppTypeByStr(field.mType);
                if (i > 0)
                {
                    construct += string.Format(", const {0}& _{1}", fixType, fname);
                }
                else
                {
                    construct += string.Format("const {0}& _{1}", fixType, fname);
                }
            }
            construct += ")\r\n\t:";
            for (int i = 0; i < excelFields.Count; ++i)
            {
                ExcelField field = excelFields[i];
                if (field.skip_export_bin)
                {
                    continue;
                }

                string fname = field.name.ToLower();
                if (i > 0)
                {
                    construct += string.Format(", {0}(_{0})", fname);
                }
                else
                {
                    construct += string.Format("{0}(_{0})", fname);
                }
            }
            construct += "{};\r\n";

            string content = "";
            for (int i = 0; i < excelFields.Count; ++i)
            {
                ExcelField field = excelFields[i];
                if (field.skip_export_bin)
                {
                    continue;
                }

                string fname = field.name.ToLower();
                string fixType = Assist.GetCppTypeByStr(field.mType);
                content += string.Format("\t{0} {1};\n", fixType, fname);
            }

            string tail = string.Format("\tstatic const Table_{0}* Get({1});\n\tstatic const std::map<{1}, const Table_{0}*>& GetAll();\n}};\n", tableName, keyType);

            str += head + construct + content + tail;
        }



        public void AppendCSharpAccess(ref string str, string keyType, string outputFilename)
        {
            var fields = excelFields.FindAll(x => /*(x.isPrimary || x.is_bin_cs) &&*/ (!x.skip_export_bin));
            if (fields == null || fields.Count <= 0)
                return;

            StringBuilder sb = new StringBuilder();
            string className = "Table_" + tableName;
            sb.AppendFormat(@"
public sealed class {0}
{{
", className);
            sb.AppendFormat(@"
    public {0}(", className);
            for (int i = 0; i < fields.Count; ++i)
            {
                ExcelField field = fields[i];
                string fname = field.name.ToLower();
                string fixType = Assist.GetCSharpTypeByStr(field.mType);
                if (i > 0)
                {
                    sb.Append(@", ");
                }
                sb.AppendFormat(@"{0} _{1}", fixType, fname);
            }
            sb.Append(@")
    {");
            for (int i = 0; i < fields.Count; ++i)
            {
                ExcelField field = fields[i];
                string fname = field.name.ToLower();
                sb.AppendFormat(@"
        {0} = _{0};", fname);
            }
            sb.Append(@"
    }
");

            for (int i = 0; i < fields.Count; ++i)
            {
                ExcelField field = fields[i];

                string fname = field.name.ToLower();
                string fixType = Assist.GetCSharpTypeByStr(field.mType);
                sb.AppendFormat(@"
    public {0} {1};", fixType, fname);
            }


            sb.AppendFormat(@"

    public static List<{1}> DataList 
    {{ 
        get 
        {{
            return __data.Values.ToList();
        }} 
    }}

    private static Dictionary<{0}, {1}> __data = new Dictionary<{0}, {1}>();", keyType, className);

            sb.AppendFormat(@"
    public static {0} Get({1} __v)
    {{
        {0} __r;

        if (__data.TryGetValue(__v, out __r)) 
            return __r;

        return null;
    }}", className, keyType);

            sb.Append(@"
    public static void Load(string dirName, VFSPackage vfsPackage, DebugModeType modeType)
    {");
            //sb.Append(FormatCSharpAccess(outputFilename));

            string binFilename = (Path.GetFileNameWithoutExtension(outputFilename) + ".bytes").ToLower();
            sb.AppendFormat(@"
        var bytes = Exiledgirls.Frameworks.Helper.FileHelper.LoadFileStream(""{0}"", dirName, vfsPackage, modeType);", binFilename);
            sb.Append(@"
        if (bytes == null)
            return;

        using (BinaryReader br = new BinaryReader(new MemoryStream(bytes)))
        {
            int head = br.ReadInt32();
            if (head != 87654321) 
                return;

            int __count = br.ReadInt32();
            for(int i = 0; i < __count; ++i)
            {");
            for (int i = 0; i < fields.Count; ++i)
            {
                var field = fields[i];
                sb.AppendFormat(@"
                {0} {1} = {2}", Assist.GetCSharpTypeByStr(field.mType),
                    field.name.ToLower(), Assist.GetCSRtString(field.mType));
            }
            sb.AppendFormat(@"
                __data[{0}] = new Table_{1} (", GetPramaryKeyName(), tableName);

            for (int i = 0; i < fields.Count; ++i)
            {
                var field = fields[i];
                if (i != 0)
                {
                    sb.Append(@", ");
                }

                sb.Append(field.name.ToLower());
            }
            sb.Append(@");");
            sb.Append(@"
            }
        }
    }
}");
            str += sb.ToString();
        }

        public void AppendGoDefine(ref string str, string outputFilename)
        {
            string head = string.Format("type {0} struct {{\n\tfile string\n", tableName);
            string content = "";
            string line = "";
            string newfunc = string.Format("func new{0} (", tableName);
            string newreturn = string.Format("\treturn {0}{{\"{1}\",", tableName, outputFilename);
            for (int i = 0; i < excelFields.Count; ++i)
            {
                ExcelField field = excelFields[i];
                ExcelField fieldNext = null;
                if (i != excelFields.Count - 1)
                {
                    fieldNext = excelFields[i + 1];
                }

                if (!field.client_only)
                {
                    if (i != excelFields.Count - 1)
                    {
                        newfunc += string.Format("{0} {1},", field.name, TypeToGo(field.mType));
                        newreturn += string.Format("{0},", field.name);
                    }
                    else
                    {
                        newfunc += string.Format("{0} {1}) {2}{{\n", field.name, TypeToGo(field.mType), tableName);
                        newreturn += string.Format("{0}}}\n}}\n\n\n", field.name);
                    }

                    line = string.Format("\t{0} {1}\n", field.name, TypeToGo(field.mType));
                    content += line;
                }
            }
            str += (head + content + "}\n" + newfunc + newreturn);

        }

        public bool AppendGoDefine2(ref string str, string outputFilename)
        {
            string head = string.Format("\ntype {0} struct {{\n", tableName);
            string content = "";
            string line = "";
            string fieldType = "";
            if (export_golang.Length != 0 && export_golang == "0")
            {
                return false;
            }
            for (int i = 0; i < excelFields.Count; ++i)
            {
                ExcelField field = excelFields[i];
                ExcelField fieldNext = null;
                if (i != excelFields.Count - 1)
                {
                    fieldNext = excelFields[i + 1];
                }

                if (!field.client_only)
                {

                    line = string.Format("\t{0} {1}\n", field.name, TypeToGo(field.mType));
                    content += line;
                }

                if (field.isPrimary)
                {
                    if (!field.mType.Equals("string"))
                    {
                        fieldType = "int64";
                    }
                    else
                    {
                        fieldType = "string";
                    }
                }

            }
            str = (head + content + "}\n" + "var " + tableName + "Map = map[" + fieldType + "] *" + tableName + "{\n");


            return true;
        }


        public void AppendErlangDefine(ref string str, string outputFilename)
        {
            if (export_erl.Length != 0 && export_erl == "0")
            {
                return;
            }
            string head = string.Format("-record(res_{0},{{\n\tfile_name=\"{1}\" :: string(),\n", tableName.ToLower(), outputFilename);
            string content = "";
            string line = "";
            for (int i = 0; i < excelFields.Count; ++i)
            {
                ExcelField field = excelFields[i];
                ExcelField fieldNext = null;
                if (i != excelFields.Count - 1)
                {
                    fieldNext = excelFields[i + 1];
                }


                if (!field.client_only)
                {
                    line = string.Format("\t{0} :: {1}", field.name.ToLower(), TypeToErl(field.mType));

                    if (i != excelFields.Count - 1)
                    {
                        line += ",\n";
                    }
                    else
                    {
                        line += "\n";
                    }
                    content += line;
                }
                else
                {
                    if (i == excelFields.Count - 1)
                    {
                        content = content.Substring(0, content.LastIndexOf(",")) + "\n";
                    }
                }
            }
            str += (head + content + "}).\n\n");
        }

        public void AppendGoFunc(ref string str, string outputFilename)
        {
            string head = string.Format("\tload_{0}_csv(Path)\n", Path.GetFileNameWithoutExtension(outputFilename));
            str += head;
        }

        public void AppendGoDeclaration(ref string str, string outpuFilename)
        {
            string type = "";
            for (int i = 0; i < excelFields.Count; ++i)
            {
                ExcelField field = excelFields[i];
                if (field.isPrimary)
                {

                    type = TypeToGo(field.mType);
                    break;
                }
            }
            string head = string.Format("\t{0}Map map[{1}] {2}\n", tableName, type, tableName);
            str += head;

        }

        public void AppendGoImpl(ref string str, string outputFilename)
        {
            string funStr = "";
            string mainKey = "";
            string type = "";
            string initFunc = "";
            initFunc += string.Format("new{0}(", tableName);
            for (int i = 0; i < excelFields.Count; ++i)
            {
                ExcelField field = excelFields[i];
                if (field.isPrimary)
                {
                    type = TypeToGo(field.mType);
                    mainKey = field.name;
                    break;
                }
            }

            for (int i = 0; i < excelFields.Count; ++i)
            {
                //Console.Write(i);
                ExcelField field = excelFields[i];
                if (i != excelFields.Count - 1)
                {
                    funStr = string.Format("{0}(v[{1}]), ", TypeToFun(field.mType), i);
                    initFunc += funStr;
                }
                else
                {
                    funStr = string.Format("{0}(v[{1}]))", TypeToFun(field.mType), i);
                    initFunc += funStr;
                }
            }

            string head = string.Format("func load_{0}_csv(Path string) (err error){{ \n", Path.GetFileNameWithoutExtension(outputFilename));
            head += string.Format("\t{0}Map = make(map[{1}] {2})\n", tableName, type, tableName);
            head += string.Format("\tfilePath := filepath.Join(Path, \"{0}\")\n", outputFilename);
            head += "\tdata, err := read_csv_line(filePath)\n";
            head += "\tif err != nil {\n";
            head += "\t\treturn err\n";
            head += "\t}\n\n";
            head += "\tfor _, v := range data {\n";
            head += string.Format("\t\tvalue := {0}\n", initFunc);
            head += string.Format("\t\t{0}Map[value.{1}] = value\n", tableName, mainKey);
            head += "\t}\n";
            head += "\treturn nil\n";
            head += "}\n\n";

            str += head;

        }

        public void AppendErlangImpl(ref string str, string outputFilename)
        {
            if (export_erl.Length != 0 && export_erl == "0")
            {
                return;
            }
            string mainKey = "";
            for (int j = 0; j < excelFields.Count; ++j)
            {
                ExcelField field = excelFields[j];
                if (field.isPrimary)
                {
                    mainKey = field.name.ToLower();
                    break;
                }
            }

            string head = string.Format("load_csv(\"{0}\", Path) ->\n", outputFilename);
            head += string.Format("\tFilePath =  Path ++ \"{0}\",\n", outputFilename);
            head += "\t{ok, Data} = read_file(FilePath),\n";
            head += "\tLineList = read_csv_line(Data),\n";
            head += string.Format("\tets:new(res_{0}, [set, protected, named_table, {{keypos, #res_{1}.{2}}}]),\n", tableName.ToLower(), tableName.ToLower(), mainKey);
            head += "\tlists:map(\n";
            head += "\t\tfun(Line) ->\n";
            head += string.Format("\t\t\tValue = \n\t\t\t\t#res_{0} {{\n", tableName.ToLower());

            string content = "";
            for (int i = 0, count = 1; i < excelFields.Count; ++i)
            {
                ExcelField field = excelFields[i];

                if (!field.client_only)
                {
                    string line = string.Format("\t\t\t\t\t{0} = {1}(lists:nth({2}, Line))", field.name.ToLower(), TypeToFun(field.mType), count);

                    if (i != excelFields.Count - 1)
                    {
                        line += ",\n";
                    }
                    else
                    {
                        line += "\n\t\t\t\t},\n";
                    }
                    content += line;
                    count++;
                }
                else
                {
                    if (i == excelFields.Count - 1)
                    {
                        content = content.Substring(0, content.LastIndexOf(",")) + "\n\t\t\t\t},\n"; ;
                    }
                }
            }

            content += string.Format("\t\t\tets:insert(res_{0}, Value)\n", tableName.ToLower());

            str += (head + content + "\t\tend, LineList);\n\n\n");
        }



        public void AppendCSharpDefine(ref string str, string outputFilename)
        {
            if (export_csharp.Length == 0 ||
                export_csharp == "0")
            {
                return;
            }

            string head = string.Format("public class TD_{0} \n{{\n    public const string _FILENAME=\"{1}\";\n",
                tableName, outputFilename);

            string content = "";
            int num = 0;
            foreach (ExcelField field in excelFields)
            {
                //if (field.is_bin_cs || field.isPrimary == true)
                {
                    if (content != "")
                    {
                        content += "\n";
                    }
                    content += string.Format("    public const int {0} = {1};", field.name, field.index);
                    num = num + 1;
                }
            }
            string tail = "\n}\n";
            if (num > 1)
            {
                if (tableDesc.Length > 0)
                {
                    str += string.Format("//{0}\n", tableDesc);
                }

                str += head + content + tail;
            }

        }

        public bool AppendXmlDefind(ref string str)
        {
            if (export_xml.Length == 0 || export_xml == "0")
            {
                return false;
            }
            string head = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>";
            string table = string.Format("<{0}>", tableName);
            str = string.Format("{0}\n{1}\n{2}\n</{3}>", head, table, str, tableName);
            return true;
        }

        public void AppendLuaDefine(ref string str, string outputFilename)
        {
            string head = string.Format("{0}=\n{{\nmeta=\n{{\n", tableName);

            string content = "";
            int key = 0;
            int num = 0;
            for (int i = 0; i < excelFields.Count; ++i)
            {
                ExcelField field = excelFields[i];

                //if (field.is_bin_lua || field.isPrimary == true)
                {
                    if (field.isPrimary)
                    {
                        key = i + 1;
                    }

                    string typeFlag = string.Empty;
                    if (field.mType == "string")
                    {
                        typeFlag = "s";
                    }
                    else if (field.mType == "double")
                    {
                        typeFlag = "f";
                    }
                    else if (field.mType.Equals("centimeter") || field.mType.Equals("decimeter") || field.mType.Equals("ratio") || field.mType.Equals("millimetre"))
                    {
                        typeFlag = "f";
                    }
                    else
                    {
                        typeFlag = "i";
                    }

                    string line = string.Format(
                        "{{\"{0}\",\"{1}\"}},\n",
                        field.name.ToLower(), typeFlag);

                    content += line;
                    num = num + 1;
                }
            }

            int extPos = outputFilename.LastIndexOf('.');
            outputFilename = outputFilename.Substring(0, extPos);

            string tail = string.Format("}},\nkey={0},\nsrc=\"{1}\",\n}},\n\n", key, outputFilename);
            if (num > 1)
            {
                str += (head + content + tail);
            }
        }

    }
}