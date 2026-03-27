using System;
using System.Collections.Generic;
using System.Text;


namespace ExcelTool
{
    public partial class ConvertTool
    {
        private string FormatXmlString(List<CellDataForLua> input , string tableName)
        {
            StringBuilder content = new StringBuilder();
            string key = string.Empty;

            int fieldCount = fieldConfig.excelFields.Count;

            for (int i = 0; i < fieldCount; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                CellDataForLua cellData = input[i];

                //不跳过当前单元格
                string cellString = string.Empty;
                if (!field.client_only)
                {
                    cellString = string.Format(" {0}=\"{1}\"", field.name, cellData.GetOrginalString());
                    content.Append(cellString);
                }
            }

            return string.Format("\t<data{0} />\r\n", content.ToString());
        }
    }
}

