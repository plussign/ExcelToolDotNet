using System.IO;
using System.Text;


namespace ExcelTool
{
    public partial class ConvertTool
    {
        private void SaveCppFile()
        {
            StringBuilder sb = new StringBuilder();

            sb.Append("#include \"stdafx.h\"\n");
            sb.Append("#include \"HeadImport.h\"\n");

            for (int i = 0; i < allCppContent.Count; ++i)
            {
                sb.Append(allCppContent[i]);
            }

            sb.Append("bool __LoadDataTable__()\n{\n");
            for (int i = 0; i < allTableName.Count; ++i)
            {
                string err = "ZBError(\"LoadTableFail:" + allTableName[i] + "\");";
                string line = string.Format("\tif (!_InitTable_{0}())\n\t{{\n\t\t{1}\n\t\treturn false;\n\t}}\n", allTableName[i], err);
                sb.Append(line);
            }
            sb.Append("\n\treturn true;\n}\n");
            sb.Append("void _FreeDataTable_()\n{");
            for (int i = 0; i < allTableName.Count; ++i)
            {
                string line = string.Format("\t_FreeTable_{0}();\n", allTableName[i]);
                sb.Append(line);
            }
            sb.Append("}");


            BaseHelper.WriteText("DATA_TABLE_LOAD.cpp", sb.ToString());
        }

        private string GetAllCppContent()
        {
            string cppPrimaryKey = fieldConfig.GetCppPrimaryKey();
            string cppMapTypeName = string.Format("std::map<{0},const Table_{1}*>", cppPrimaryKey, fieldConfig.tableName);
            string binFilename = ( Path.GetFileNameWithoutExtension(output.filename) + ".bin" ).ToLower();

            string cppContent = string.Empty;
            cppContent += string.Format("\nstatic {3} data_{2};\nstatic bool _InitTable_{2}() \n{{\n{1}\n}}\n",
            cppPrimaryKey, FormatBinReadText(binFilename), fieldConfig.tableName, cppMapTypeName);

            cppContent += string.Format("static void _FreeTable_{0}()\n{{\n\tfor({1}::const_iterator i = data_{0}.begin(); " +
                "i!= data_{0}.end(); ++i)\n\t{{\n\t\tdelete i->second;\n\t}}\n\tdata_{0}.clear();\n}}\n",
                fieldConfig.tableName, cppMapTypeName);

            cppContent += string.Format("const Table_{0}* Table_{0}::Get({1} v)\n{{\n\t{2}::const_iterator i = data_{0}.find(v); return i==data_{0}.end() ? NULL : i->second;\n}}\n"
                , fieldConfig.tableName, cppPrimaryKey, cppMapTypeName);

            cppContent += string.Format("const {0}& Table_{1}::GetAll()\n{{\n\treturn data_{1};\n}}\n", cppMapTypeName, fieldConfig.tableName);
            return cppContent;
        }
    }
}

