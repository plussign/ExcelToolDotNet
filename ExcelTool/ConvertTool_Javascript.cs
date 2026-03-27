using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    public partial class ConvertTool
    {
        public string _fmtstr(string s)
        {
            s = s.Replace("\"", "\\\"");
            return "\"" + s + "\"";
        }

        public string formatJsLine(List<CellDataForLua> input)
        {
            string output = string.Empty;
            string primary = string.Empty;

            for (int i = 0; i < fieldConfig.excelFields.Count; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                string s = string.Empty;

                if (field.isPrimary)
                {
                    primary = _fmtstr(input[i].GetOrginalString()) + ":";
                }

                if (field.mType == "string")
                {
                    s = _fmtstr(input[i].GetOrginalString());
                }
                else if (field.mType == "int" || field.mType == "double")
                {
                    s = input[i].GetOrginalString();
                }
                else
                {
                    s = _fmtstr(input[i].GetOrginalString());
                }

                if (i > 0)
                {
                    output += ("," + s);
                }
                else
                {
                    output = s;
                }
            }

            return primary + "[" + output + "]";
        }

        public void SaveAllJsDefine()
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < allJsContent.Count; ++i)
            {
                sb.Append(allJsContent[i]);
            }

            StringBuilder exports = new StringBuilder();
            for (int i = 0; i < allJsHead.Count; ++i)
            {
                string s = string.Format("\"{0}\": Table_{0},\n", allJsHead[i]);
                exports.Append(s);
            }

            string tail = "module.exports={" + exports + "};";

            BaseHelper.WriteTextNoBOM("DataTableHead.js", sb.ToString() + tail);
        }

        private string GetAllJsDefine()
        {
            StringBuilder sb = new StringBuilder();

            string head = string.Format("class Table_{0}{{\nconstructor(dt){{\nthis.data= dt;\n}};\n_assign(dt){{\nthis.data= dt;\n}};\n", fieldConfig.tableName);
            sb.Append(head);

            uint slotIndex = 0;
            for (int i = 0; i < fieldConfig.excelFields.Count; ++i)
            {
                string s = string.Format("get {0}() {{ return this.data[{1}]; }}; \n", fieldConfig.excelFields[i].name, slotIndex);
                sb.Append(s);
                slotIndex++;
            }

            sb.Append("};\n");

            return sb.ToString();
        }
    }
}
