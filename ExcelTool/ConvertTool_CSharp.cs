using System.Collections.Generic;
using System.Text;


namespace ExcelTool
{
    public partial class ConvertTool
    {
        private void SaveCSharpAccessInterface()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(@"
using System.Collections.Generic;
using Exiledgirls.Frameworks.Define;
using System.IO;
using System.Linq;
using Mayday.Frameworks.Helper;
using Mayday.Fixed.Frameworks;
");
            sb.Append(CSharpAccessIntefaceDefine);

            sb.Append(@"

public sealed class DataTableAllTabler
{
    public static void Load(string dirName, VFSPackage vfsPackage, DebugModeType modeType)
    {");
            for (int i = 0; i < allTableName.Count; ++i)
            {
                sb.AppendFormat(@"
        Table_{0}.Load(dirName, vfsPackage, modeType);", allTableName[i]);
            }
            sb.Append(@"
    }
}");

            BaseHelper.WriteText("TABLE_FIELD_ACCESS.cs", sb.ToString());
        }

        private string FormatCSVLineString(List<string> content)
        {
            if (content.Count == 0)
            {
                return "";
            }

            StringBuilder sb = new StringBuilder();

            sb.Append(content[0]);

            for (int i = 1; i < content.Count; ++i)
            {
                sb.Append(Def.CELL);
                sb.Append(content[i]);
            }

            return sb.ToString();
        }
    }
}
