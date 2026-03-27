using System.Collections.Generic;
using System.Text;

public class CustomEnumMgr
{
    static public bool Add(string name, int value)
    {
        enums.Add(new KeyValuePair<string, int>(name, value));
        return true;
    }

    static public string ToErlang()
    {
        StringBuilder sb = new StringBuilder();

        foreach (var kv in enums)
        {
            string str = string.Format("-define({0}, {1}).\n", kv.Key, kv.Value);
            sb.Append(str);
        }

        return sb.ToString();
    }

    static public string ToLua()
    {
        StringBuilder sb = new StringBuilder();

        foreach (var kv in enums)
        {
            string str = string.Format("{0}={1}\n", kv.Key, kv.Value);
            sb.Append(str);
        }

        return sb.ToString();
    }

    static public string ToCpp()
    {
        StringBuilder sb = new StringBuilder();

        foreach (var kv in enums)
        {
            string str = string.Format("    {0}={1},\r\n", kv.Key, kv.Value);
            sb.Append(str);
        }

        return sb.ToString();
    }

    static public string ToCSharp()
    {
        StringBuilder sb = new StringBuilder();

        foreach (var kv in enums)
        {
            string str = string.Format("    {0}={1},\r\n", kv.Key, kv.Value);
            sb.Append(str);
        }

        return sb.ToString();
    }

    static public string ToGo()
    {
        
        StringBuilder sb = new StringBuilder();

        sb.Append("const (\r\n");
        foreach (var kv in enums)
        {
            string str = string.Format("    {0} = {1}\r\n", kv.Key, kv.Value);
            sb.Append(str);
        }
        sb.Append(")\r\n");

        return sb.ToString();
    }

    public static List<KeyValuePair<string, int>> Enums { get { return enums; } }

    private static List<KeyValuePair<string, int>> enums = new List<KeyValuePair<string, int>>();
}
