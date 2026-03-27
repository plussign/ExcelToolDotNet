using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ExcelTool
{
    public partial class ConvertTool
    {
        //static Random r = new Random(System.Environment.TickCount);
        private string FormatGoLineString(List<CellDataForLua> input , string tableName)
        {
            StringBuilder content = new StringBuilder();
            string key = string.Empty;

            uint[] skippedFields = null;
            int fieldCount = fieldConfig.excelFields.Count;
            if (fieldCount > 2)
            {
                int skippedFieldsRecords = (int)Math.Ceiling((double)fieldCount / 32);
                skippedFields = new uint[skippedFieldsRecords];
                for (int i = 0; i < skippedFieldsRecords; ++i)
                {
                    skippedFields[i] = 0;
                }
            }

            for (int i = 0; i < fieldCount; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                CellDataForLua cellData = input[i];

                bool isStringField = field.mType.Equals("string");
                bool isSkippableKey = (field.isPrimary && !isStringField);

                //不跳过当前单元格
                string cellString = string.Empty;
                if (!field.client_only)
                {
                    if (isStringField)
                    {
                        cellString = Assist.ToGoStr(cellData.GetOrginalString());
                    }
                    else
                    {
                        cellString = cellData.GetOrginalString();
                    }

                    if (content.Length > 0)
                    {
                        content.Append(",");
                    }
                    content.Append(cellString);


                    if (field.isPrimary)
                    {
                        if (!field.mType.Equals("string"))
                        {
                            key = cellData.GetOrginalString();
                        }
                        else
                        {
                            key = Assist.ToGoStr(cellData.GetOrginalString());
                        }
                    }
                }
            }
            //StringBuilder sbSkippedRecords = new StringBuilder();
            //if (null != skippedFields)
            //{
            //    for (int i = 0; i < skippedFields.Length; ++i)
            //    {
            //        sbSkippedRecords.Append(skippedFields[i].ToString());
            //        sbSkippedRecords.Append(",");
            //    }
            //}

            return string.Format("\t{0}:&{1}{{{2}}},\r\n", key, tableName , content.ToString());
            //return string.Format("[{0}]=\"{1}\",\r\n", key, r.Next(1000, 9999));
        }
    }
}

/*
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool
{
    public partial class ConvertTool
    {

        const string csv_go_begin = @"package csv
                         
import (
    ""io/ioutil""
    ""strings""
    ""path/filepath""
    ""strconv""
)

func read_csv_line(file string) (value [][]string, err error) {
    fileData, err := ioutil.ReadFile(file)
    if err!= nil {
        return nil, err
    }
    strData := string(fileData)
    newData := strings.Split(strData, ""/b"")
    data := make([][]string, len(newData))
    for k, v := range newData {
        newData1 := strings.Split(v, ""/c"")
        data[k] = newData1
    }
    return data, nil
}

func to_list(value string) string {
    return value
}

func to_float(value string) float64 {
    f, err := strconv.ParseFloat(value, 64)
    if err != nil {
        return float64(0)
    } else {
        return float64(f)
    }
}

func to_int(value string) int64 {
    i, err := strconv.Atoi(value)
    if err != nil {
        return int64(0)
    } else {
        return int64(i)
    }
}



";

    }
}
*/
