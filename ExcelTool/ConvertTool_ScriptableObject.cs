using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    public partial class ConvertTool
    {
        private string FormatScriptableObject(List<CellDataForLua> input)
        {
            StringBuilder content = new StringBuilder();
            string key = string.Empty;
            string keyValue = string.Empty;

            int fieldCount = fieldConfig.excelFields.Count;

            for (int i = 0; i < fieldCount; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                CellDataForLua cellData = input[i];

                bool isStringField = field.mType.Equals("string");
                string cellString = string.Empty;

                if (isStringField)
                {
                    if (CellDataForLua.CellTypeForLua.Standard == cellData.type)
                    {
                        uint textIndex = I18N.RegisterText(cellData.GetOrginalString(), false);
                        if (textIndex > 0)
                        {
                            cellString = textIndex.ToString();
                        }
                        else
                        {
                            cellString = "0";
                        }
                    }
                    else
                    {
                        //索引词条表
                        uint strId = cellData.GetStringIndex();
                        cellString = strId.ToString();
                    }
                }
                else
                {
                    //int double 或者枚举值类型的单元格
                    if (!cellData.IsBlank)
                    {
                        cellString = cellData.GetOrginalString();
                    }
                    else
                    {
                        cellString = "0";
                    }
                }


                if (i == 0)
                {
                    content.Append("  - ");
                }
                else
                {
                    content.Append("    ");
                }

                content.Append(field.name.ToLower());
                content.Append(": ");
                content.Append(cellString);
                content.AppendLine();
            }


            return content.ToString();
        }
    }
}
