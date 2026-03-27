using System;
using System.Collections.Generic;
using System.Text;


namespace ExcelTool
{
    public partial class ConvertTool
    {
        private string FormatHumanReadable(List<CellDataForLua> input)
        {
            StringBuilder content = new StringBuilder();
            string key = string.Empty;
            string keyValue = string.Empty;
            
            int fieldCount = fieldConfig.excelFields.Count;

            for (int i = 0; i < fieldCount; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                CellDataForLua cellData = input[i];
                
                string cellString = cellData.GetSingleRowString();

                content.Append(field.name);
                content.Append("[");
                content.Append(cellString);
                content.Append("]\t");

                if (field.isPrimary)
                {
                    key = field.name;
                    keyValue = cellString;
                }
            }
            
            return string.Format("[KEY]{0}[{1}]\t{2}\r\n", key, keyValue, content.ToString());
        }
    }
}
