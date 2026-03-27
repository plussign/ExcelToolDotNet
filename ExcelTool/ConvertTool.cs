using ExcelTool.ExcelTool;
using libxl;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace ExcelTool
{
    public partial class ConvertTool
    {
        private InputConfig input;
        private OutputConfig output;
        private EnumManager enumList;
        private FieldConfig fieldConfig;
        private List<string> primarys;
        private EnumManager checkEnumList;
        private Dictionary<string, string> checkHints = new Dictionary<string, string>();
        private Dictionary<string /* 缓存的表 */ , Dictionary<string /* 缓存的列名*/, Dictionary<int,string> /*某一列的集合 */ >> allLoadCfgData =
            new Dictionary<string, Dictionary<string, Dictionary<int, string>>>();

        private Dictionary<int, string> tmpCheckList = null;

        private bool passCheck = true;

        private string luaDefine = "";
        private string luaEnumText = "";
        private string heroPropDefine = "";
        private string goEnumText = "";

        private string cppDefine = "";
        private string goDefine = "";
        private string csharpDefine = "";
        private string CSharpAccessIntefaceDefine = "";

        private string erlDefine = "";
        private string erlImpl = "";
        private string erlImplEnd = "";

        private Dictionary<string, string> hints = new Dictionary<string, string>();
        private List<string> allTableName = new List<string>();
        private List<string> allCppContent = new List<string>();

        private List<string> allJsContent = new List<string>();
        private List<string> allJsHead = new List<string>();

        private List<string> allOutLuaName = new List<string>();
        private List<string> allOutCSharpName = new List<string>();

        private StringBuilder InOutMapString = new StringBuilder();


        public ConvertTool()
        {
            XmlDocument docEnums = new XmlDocument();
            try
            {
                string path = Path.Combine("config/enums.xml");
                if (Program.isDynamicOutPut)
                {
                    string dynamicPath = Path.Combine("../Config/config/enums.xml");
                    if (File.Exists(dynamicPath))
                    {
                        path = dynamicPath;
                    }
                }
                string xmlContent = File.ReadAllText(path, Encoding.UTF8);
               // GlobeError.Push(xmlContent);
                docEnums.LoadXml(xmlContent);
            }
            catch( System.Exception e)
            {
                GlobeError.Push(string.Format("枚举文件载入失败 enums.xml={0}", e.ToString()));
                return;
            }
            if (docEnums.DocumentElement == null)
            {
                GlobeError.Push("枚举文件载入失败 enums.xml");
                return;
            }
            enumList = new EnumManager();
            checkEnumList = new EnumManager();
            foreach (XmlElement child in docEnums.DocumentElement)
            {
                enumList.Load(child);
                checkEnumList.Load(child);
                erlDefine += enumList.ExportErlCode();
                erlDefine += "\n";
            }
        }

        //private string GetHeroPorpEnum()
        //{
        //    string content = "";
        //    if (enumList.items.ContainsKey("PropertyType"))
        //    {
        //        Dictionary<string, EnumItem> enumItems = enumList.items["PropertyType"];
        //        foreach (string key in enumItems.Keys)
        //        {
        //            content += string.Format("[{0}] = \"{1}\",\n", enumItems[key].luaName, key);
        //        }
        //    }
        //    else
        //    {
        //        GlobeError.Push(string.Format("无法找到枚举类型:[{0}]!", "PropertyType"));
        //    }
        //    return content;
        //}
        
        private bool ConvertTableContents(
            ref string humanReadable,
            ref string scriptableObject,
            ref string csv,
            ref string go,
            ref string lua,
            ref string cpp_, 
            ref byte[] bin_, 
            ref byte[] csharpBin,
            ref byte[] luaBin,
            ref string js,
            ref string luaDynamic,
            ref string xml)
        {
            hints.Clear();
            
            StringBuilder allHumanReadableLine = new StringBuilder();
            StringBuilder allCsvLine    = new StringBuilder();

            StringBuilder allLua = new StringBuilder();

            StringBuilder allLuaDynamic = new StringBuilder();

			StringBuilder allXml = new StringBuilder();

            StringBuilder allGo = new StringBuilder();
            StringBuilder allJs = new StringBuilder();

            List<byte[]> allBinLua = new List<byte[]>();
            List<byte[]> allBinCSharp = new List<byte[]>();

            StringBuilder allCpp = new StringBuilder();
            List<byte[]> allBin = new List<byte[]>();

            StringBuilder allScriptableObject = new StringBuilder();

            int fieldCount = fieldConfig.excelFields.Count;

            //判断是否主键类型是字符串【客户端LuaBin格式用】
            bool iskeyStr = false;
            for (int i = 0; i < fieldCount; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                if(field.isPrimary)
                {
                    if (!field.mType.Equals("string"))
                    {
                        iskeyStr = false;
                    }
                    else
                    {
                        iskeyStr = true;
                    }
                    break;
                }
            }

            foreach (InputConfig.SourceFileInfo fileInfo in input.files)
            {
                string file = fileInfo.fileName;      
                    
                string postfix = Path.GetExtension(file).ToLower();
                string oldDirname = Path.GetDirectoryName(file);
                if (!File.Exists(file))
                {
                    if (postfix == ".xls")
                    {
                        string plainFileName = Path.Combine(oldDirname, Path.GetFileNameWithoutExtension(file));
                        Console.WriteLine($"\n[ConvertTableContents]文件[{plainFileName}]不存在，尝试替换xlsx缀名后重新寻找...");
                        if (File.Exists(plainFileName + ".xlsx"))
                        {
                            file = plainFileName + ".xlsx";
                        }
                    }
                    else if (postfix == ".xlsx")
                    {
                        string plainFileName = Path.Combine(oldDirname, Path.GetFileNameWithoutExtension(file));
                        Console.WriteLine($"\n[ConvertTableContents]文件[{plainFileName}]不存在，尝试替换xls后缀名后重新寻找...");
                        if (File.Exists(plainFileName + ".xls"))
                        {
                            file = plainFileName + ".xls";
                        }
                    }
                }
                
                SheetCache sheet = SheetCacheMgr.GetCache(file);
                if (sheet == null) 
                {
                    Log.Write("==>>Open[{0}]...", file);

                    if (!File.Exists(file))
                    {
                        if (true == fileInfo.sourceDynamic)
                        {
                            continue; 
                        }
                        GlobeError.Push(string.Format("excel文件不存在:{0}", file));
                        return false;
                    }


                    Book book = XlsLoader.LoadBook(file);
                    if (book == null)
                    {
                        return false;
                    }
                    sheet = new SheetCache(book.getSheet(0));
                    
                    book.Dispose();
                    book = null;

                    SheetCacheMgr.AddExcelFileCache(file, sheet);

                    Log.WriteLine("Done");
                }
                else
                {
                    Log.WriteLine("$$>>ReUse[{0}]...", file);
                }
                
                if (sheet == null)
                {
                    GlobeError.Push(string.Format("无法获得excel工作薄:{0}", file));
                    return false;
                }

                if (!fieldConfig.LoadSlotInfo(sheet, file))
                {
                    return false;
                }

                int height = sheet.lastRow();
                for (int i = 1; i < height; ++i)
                {
                    string key = string.Empty;
                   
                    List<CellDataForLua> clientSingleRawLine = new List<CellDataForLua>();
                    List<string> serverSingleRawLine = new List<string>();

                    if (!ReadExcelRawLine(file, sheet, i, ref key, ref clientSingleRawLine, ref serverSingleRawLine))
                    {
                        return false;
                    }

                    if (Program.i18nExtraOnly)
                    {
                        continue;
                    }

                    string humanReadableLine = FormatHumanReadable(clientSingleRawLine);
                    allHumanReadableLine.Append(humanReadableLine);

                    string scriptableObjectLine = FormatScriptableObject(clientSingleRawLine);
                    allScriptableObject.Append(scriptableObjectLine);

                    string luaLineDyanmic = FormatLuaLineStringDynamic(clientSingleRawLine);
                    allLuaDynamic.Append(luaLineDyanmic);

                    string goLine = FormatGoLineString(clientSingleRawLine, fieldConfig.tableName);
                    allGo.Append(goLine);

                    string xmlLine = FormatXmlString(clientSingleRawLine, fieldConfig.tableName);
                    allXml.Append(xmlLine);

                    byte[] lineBinLua = FormatBinLuaLineString(clientSingleRawLine, iskeyStr);
                    if (lineBinLua != null)
                    {
                        allBinLua.Add(lineBinLua);
                    }

                    string jstemp = formatJsLine(clientSingleRawLine);

                    if (allJs.Length > 0)
                    {
                        allJs.Append(",\n");
                    }
                    allJs.Append(jstemp);
                    
                    if (fieldConfig.export_enum_only)
                    {
                        if (!ReadGlobeEnum(clientSingleRawLine))
                        {
                            GlobeError.Push(string.Format("解析专用枚举失败"));
                            return false;
                        }
                    }
                }

                GC.Collect();
            }

            humanReadable = allHumanReadableLine.ToString();
            scriptableObject = allScriptableObject.ToString();
            go = allGo.ToString();
            js = allJs.ToString();
            luaDynamic = allLuaDynamic.ToString();
            xml = allXml.ToString();

            byte[] head = BitConverter.GetBytes((int)allBin.Count);
            byte[] flag = BitConverter.GetBytes((int)12345678);

            allBin.Insert(0, flag);
            allBin.Insert(1, head);

            byte[] name = UTF8Encoding.Default.GetBytes(fieldConfig.tableName);
            if ((int)allBinLua.Count != 0)
            {
                head = BitConverter.GetBytes((int)allBinLua.Count);
                allBinLua.Insert(0, BitConverter.GetBytes((int)87654321));
                allBinLua.Insert(1, BitConverter.GetBytes(iskeyStr ? (int)1 : (int)2));
                allBinLua.Insert(2, BitConverter.GetBytes(name.Length));
                allBinLua.Insert(3, name);
                allBinLua.Insert(4, head);
                luaBin = BaseHelper.Meger(allBinLua);
            }
            
            return true;
        }

        public string GetPramaryKeyName()
        {
            return fieldConfig.GetPramaryKeyName();
        }
        
        public bool ReadGlobeEnum(List<CellDataForLua> input)
        {
            string key = string.Empty;
            string value = string.Empty;

            for (int i = 0; i < fieldConfig.excelFields.Count && i < input.Count; ++i)
            {
                ExcelField field = fieldConfig.excelFields[i];
                if (field.isPrimary)
                {
                    key = input[i].GetOrginalString();
                }
                else if (field.enum_value)
                {
                    value = input[i].GetOrginalString();
                }
            }

            if (key != string.Empty && value != string.Empty)
            {
                if (int.TryParse(value, out int v))
                {
                    CustomEnumMgr.Add(key, v);
                    return true;
                }
            }

            return false;
        }

        public void BeginLoad()
        {
            erlImpl = Def.ERL_IMPL_BEGIN;

            erlDefine = "-ifndef(_CSV_HRL__).\n-define(_CSV_HRL__, true).\n\n" + erlDefine;

            luaDefine = "TableDefine=\n{\n";
            heroPropDefine = "module(\"BattleAttrMap\", package.seeall)\nStringAttr ={\n";

            if (!string.IsNullOrEmpty(Program.csv_translation_excel))
            {
                loadCsvTranslation(Program.csv_translation_excel);
            }
        }

        const string csv_erl_end = @"
load_csv(_FileName, _) -> """".

read_file(FilePath) ->
    {ok, Data} = file:read_file(FilePath),
    {ok, remove_bom(Data)}.

remove_bom(<<239, 187, 191, LastFileData/binary>>) -> LastFileData;
remove_bom(FileData) when is_binary(FileData) -> FileData.
";
        public void EndLoad()
        {
            erlImpl += Def.ERL_IMPL_END;

            luaDefine += "\n}\n";
            erlImplEnd += csv_erl_end;

            erlDefine += CustomEnumMgr.ToErlang();
            erlDefine += "\n\n-endif.";

            luaEnumText = enumList.ExportLuaCode();
            goEnumText = enumList.ExportGoCode();
            goEnumText += CustomEnumMgr.ToGo();
            //string heroPropStr = GetHeroPorpEnum();
            //heroPropDefine += heroPropStr + "}";

            BaseHelper.WriteText("TABLE_FIELD_DEFINE.lua", luaDefine);
            //BaseHelper.WriteText("TABLE_FIELD_DEFINE.h", "#pragma once\n" +  cppDefine);

            //if (Program.outputCSharpAccessInterface)
            //{
            //    SaveCSharpAccessInterface();
            //}

            SaveBinLuaList(allOutLuaName, "output_binlua_lua/list.bytes");
            //SaveBinLuaList(allOutCSharpName, "output_binlua_csharp/list.bytes");

            //BaseHelper.WriteText("TABLE_FIELD_DEFINE.cs", csharpDefine);

            BaseHelper.WriteText("GAME_ENUM_DEFINE.bytes", luaEnumText);
            BaseHelper.WriteText("GAME_ENUM_DEFINE.lua", luaEnumText);			
            //BaseHelper.WriteText("GAME_ENUM_DEFINE.h", enumList.ExportCPPCode());
            //BaseHelper.WriteText("GAME_ENUM_DEFINE.cs", enumList.ExportCSharpCode());
            BaseHelper.WriteText("GAME_ENUM_DEFINE.go", goEnumText);
            BaseHelper.WriteText("csv.hrl", erlDefine);
            
            string erlText = erlImpl + erlImplEnd;
            BaseHelper.WriteText("csv.erl", erlText);

            //BaseHelper.WriteText("BattleAttrMap.lua", heroPropDefine);

            SaveCppFile();
            SaveAllJsDefine();
            BaseHelper.WriteText("表格输入输出对应关系.txt", InOutMapString.ToString());
        }
        
        public bool Convert(string configXMLFileName)
        {
            string xmlPath = Path.Combine("config", configXMLFileName);
            if (!File.Exists(xmlPath))
            {
                GlobeError.Push("配置文件不存在 " + configXMLFileName);
                return false;
            }

            XmlDocument doc = new XmlDocument();
            try
            {
                string xmlContent = File.ReadAllText(xmlPath, Encoding.UTF8);
                doc.LoadXml(xmlContent);
            }
            catch (System.Exception e)
            {
                GlobeError.Push("配置文件XML解析失败:" + configXMLFileName + ", " + e.ToString());
                return false;
            }

            if (doc.DocumentElement == null)
            {
                GlobeError.Push("配置文件XML解析结果为空:" + configXMLFileName);
                return false;
            }

            foreach (XmlElement child in doc.DocumentElement)
            {
                OutputConfig tmpOutPut = new OutputConfig();
                tmpOutPut.Load(child);
                erlImpl += string.Format("\t\t\"{0}\",\n", tmpOutPut.filename);              // Todo 这里是csv.erl中追加的*.csv
            }

            Log.WriteLine("开始按照[{0}]转换...", configXMLFileName);

            foreach (XmlElement child in doc.DocumentElement)
            {
                if (child.Name.Equals("table"))
                {
                    ConvertTableFile(child, configXMLFileName);
                }
            }

            return true;
        }

        private string findSubFile(string rootDir, string fileName, string postfix)
        {
            if (!Directory.Exists(rootDir))
            {
                return null;
            }

            string[] allFiles = Directory.GetFiles(rootDir, "*." + postfix, SearchOption.AllDirectories);

            foreach (string prefabPath in allFiles)
            {
                if (prefabPath.ToLower().Contains(fileName.ToLower()))
                {
                    return prefabPath;
                }
            }

            return null;
        }

        private string makeScripteObjectTableHeader(string tableName)
        {
            string yamlHeader = @"%YAML 1.1
%TAG !u! tag:unity3d.com,2011:
--- !u!114 &11400000
MonoBehaviour:
  m_ObjectHideFlags: 0
  m_CorrespondingSourceObject: {fileID: 0}
  m_PrefabInstance: {fileID: 0}
  m_PrefabAsset: {fileID: 0}
  m_GameObject: {fileID: 0}
  m_Enabled: 1
  m_EditorHideFlags: 0
  m_Script: {fileID: 11500000, guid: $GUID, type: 3}
  m_Name: $TABLE_NAME
  m_EditorClassIdentifier: 
  records:
";

            yamlHeader = yamlHeader.Replace("$TABLE_NAME", tableName);

            string csMetaPath = "../Client/Assets/Scripts/DesignersTable/Define/";

            string actualMetaPath = findSubFile(csMetaPath, tableName, "cs.meta");

            /*if (string.IsNullOrEmpty(actualMetaPath))
            {
                csMetaPath = csMetaPath.Replace("Char", "LiveConcert");
                actualMetaPath = findSubFile(csMetaPath, tableName, "cs.meta");
            }*/

            if (File.Exists(actualMetaPath))
            {
                var metaText = File.ReadAllText(actualMetaPath);

                Match guidMatch = Regex.Match(metaText, @"guid:\s([a-f0-9]+)");
                string guid = guidMatch.Value.Substring(6);

                yamlHeader = yamlHeader.Replace("$GUID", guid);

                Log.WriteLine("[{0}] GUID: {1}", Path.GetFullPath(actualMetaPath), guid);
            }
            else
            {
                Log.WriteLine("XXXXXX [{0}].cs.meta NOT FOUND, should in [{1}]", tableName, actualMetaPath);
            }

            return yamlHeader;
        }

        private void ConvertTableFile(XmlElement root, string configXMLFileName)
        {
            if (!passCheck)
            {
                return;
            }

            InOutMapString.Append(string.Format("[{0}]\r\n", configXMLFileName));
            Log.Write("处理[{0}]=>", configXMLFileName);
            input = new InputConfig(Program.special_channel);
            input.Load(root);

            InOutMapString.Append(string.Format("\t输入Excel文件\r\n"));
            foreach (InputConfig.SourceFileInfo inFileInfo in input.files)
            {
                string inFile = inFileInfo.fileName;
                //if (File.Exists(inFile))
                {
                    InOutMapString.Append(string.Format("\t\t{0}\r\n", inFile));    
                }        
            }
            output = new OutputConfig();
            output.Load(root);

            string pureOutputFileName = Path.GetFileNameWithoutExtension(output.filename);
            InOutMapString.Append(string.Format("\t输出数据表\r\n\t\t{0}\r\n\r\n\r\n", pureOutputFileName));
            fieldConfig = new FieldConfig();
            fieldConfig.Load(root, configXMLFileName);

            if (Program.i18nExtraOnly)
            {
                if (!fieldConfig.hasHumanReadableText())
                {
                    Log.Write("无词条，跳过!");
                    return;
                }
                else
                {
                    Log.Write("提取需要翻译的词条...");
                }
            }
            primarys = new List<string>();

            string humanReadableContent = string.Empty;
            string scriptableObjectContent = string.Empty;
            string csvContent = string.Empty;
            string _goContent = string.Empty;
            string _luaContent = string.Empty;
            string _cppContent = string.Empty;
            string _jsContent = string.Empty;
            string _luaDynamicContent = string.Empty;
            string _xmlContent = string.Empty;

            byte[] _binContent = null;
            byte[] _binLuaCsharp = null;
            byte[] _binLua = null;

            //执行excel表格读入
            if (!ConvertTableContents(
                ref humanReadableContent, 
                ref scriptableObjectContent,
                ref csvContent,
                ref _goContent,
                ref _luaContent,
                ref _cppContent, 
                ref _binContent, 
                ref _binLuaCsharp,
                ref _binLua,
                ref _jsContent,
                ref _luaDynamicContent,
                ref _xmlContent
                ))
            {
                return;
            }

            /// 是否仅开启 导出枚举模式
            if (Program.i18nExtraOnly || fieldConfig.export_enum_only)
            {
                return;
            }

            //写入HumanReadable的表格输出文件
            string humanReadableOutFilename = string.Format("output_txt/{0}.txt", pureOutputFileName);
            string humanReadableFile = string.Format(
                "表格 [{0}] 转换后数据 \r\n\r\n{1}",
                fieldConfig.tableName, humanReadableContent);
            BaseHelper.WriteText(humanReadableOutFilename, humanReadableFile);

            //写入客户端用scriptableObject
            string scriptableObjectFilename = string.Format("output_asset/{0}.asset", pureOutputFileName);
            string scriptableObjectFile = string.Format("{0}{1}", makeScripteObjectTableHeader(fieldConfig.tableName), scriptableObjectContent);
            BaseHelper.WriteText(scriptableObjectFilename, scriptableObjectFile);


            if (Program.isDynamicOutPut)
            {
                string luaOutDynamicFilename = string.Format("output_lua_dynamic/{0}.bytes", pureOutputFileName);
                string luaDynamicContent = string.Format(
                    "TableDataSetDynamic.{0}=\r\n{{\r\n{1}}}",
                    fieldConfig.tableName, _luaDynamicContent);

                BaseHelper.WriteText(luaOutDynamicFilename, luaDynamicContent);
            }

            /*
            if (_binLuaCsharp != null)
            {
                string binLuaCSharpOutFilename = string.Format("output_binlua_csharp/{0}.bytes", pureOutputFileName);
                BaseHelper.WriteBin(binLuaCSharpOutFilename, _binLuaCsharp);
                allOutCSharpName.Add(pureOutputFileName);
            }*/

            if (_binLua != null)
            {
                string binLuaOutFilename = string.Format("output_binlua_lua/{0}.bytes", pureOutputFileName);
                BaseHelper.WriteBin(binLuaOutFilename, _binLua);
                //客户端用的二进制化的lua输出文件
                allOutLuaName.Add(pureOutputFileName);
            }


            //写入WorldServer用的cvs文件
            /*
            string csvOutFilename = string.Format("output/{0}", output.filename);
            if (fieldConfig.use_mode == TableUseMode.Common || fieldConfig.use_mode == TableUseMode.Server)
            {
                BaseHelper.WriteText(csvOutFilename, csvContent);

                fieldConfig.AppendErlangDefine(ref erlDefine, output.filename);    //  csv.hrl
                fieldConfig.AppendErlangImpl(ref erlImplEnd, output.filename);     //  csv.erl
            }
            */

            //写入go用的go文件
            string goOutFilename = string.Format("output_go/{0}.go", pureOutputFileName);
            if(fieldConfig.AppendGoDefine2(ref goDefine, output.filename))
            {
                string goContent = string.Format(
               "package {0}\n",
               "csv");

                BaseHelper.WriteText(goOutFilename, goContent + goDefine + _goContent + "}\n");
            }

            string xmlOutFilename = string.Format("output_xml/{0}.xml", pureOutputFileName);
            string xmlContent = _xmlContent;
            if (fieldConfig.AppendXmlDefind(ref xmlContent))
            {
                BaseHelper.WriteText(xmlOutFilename, xmlContent);
            }

            //写入js用的json文件
            string jsOutputFilename = string.Format("output_js/{0}.json", fieldConfig.tableName);
            BaseHelper.WriteTextNoBOM(jsOutputFilename, "{\n"+_jsContent+"\n}");


            //写入SceneServer用的bin文件，已经对应的cpp定义头
            /*
            string cppFilename = (Path.GetFileNameWithoutExtension(output.filename)).ToLower();
            string binOutFilename = string.Format("output_bin/{0}.bin", cppFilename);
            
            allCppContent.Add(GetAllCppContent());
            BaseHelper.WriteBin(binOutFilename, _binContent);
            */

            var enumCS = enumList.ExportCSharpCode();
            BaseHelper.WriteTextNoBOM("./TableEnum.cs", enumCS);


            fieldConfig.AppendLuaDefine(ref luaDefine, output.filename);
            fieldConfig.AppendCSharpDefine(ref csharpDefine, output.filename);
            fieldConfig.AppendCppDefine(ref cppDefine, fieldConfig.GetCppPrimaryKey(), output.filename);

            //if (Program.outputCSharpAccessInterface)
            {
                fieldConfig.AppendCSharpAccess(ref CSharpAccessIntefaceDefine, fieldConfig.GetCSharpPrimaryKey(), output.filename);
            }

            allJsContent.Add(GetAllJsDefine());
            allJsHead.Add(fieldConfig.tableName);
            allTableName.Add(fieldConfig.tableName);
        }       
    }
}
