using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace ExcelTool
{
    class Program
    {
        static public bool outputLog = true;
        static public bool fastConvert = false;
        static public bool i18nExtraOnly = false;
        static public bool isDynamicOutPut = false;
        static public bool outputCSharpAccessInterface = false;
        static public string special_channel = string.Empty;
        const string _special = "-special=";
        static public string csv_translation_excel = string.Empty;
        const string _csv_translation = "-csv_translation=";
        static public bool useXlsCache = false; // �Ƿ�ʹ��Cache����ת���ٶ�
        static public bool useTestData = false;

        static void ProcessCmdLine(string[] input)
        {
            List<string> args = new List<string>();
            foreach (var s in input)
            {
                args.AddRange(s.Split('`'));
            }

            foreach (var cmd in args)
            {
                string arg = cmd.ToLower();
                switch (arg)
                {
                case "-nolog":
                    {
                        outputLog = false;
                    }
                    break;

                case "-fast":
                    {
                        fastConvert = true;
                    }
                    break;

                case "-extra_text":
                    {
                        fastConvert = true;
                        i18nExtraOnly = true;
                    }
                    break;

                case "-use_xls_cache":
                    {
                        useXlsCache = true;
                    }
                    break;

                case "-use_test_data":
                    {
                        useTestData = true;
                    }
                    break;

                case "-csharp":
                    {
                        outputCSharpAccessInterface = true;
                    }
                    break;
                case "-dynamic_output":
                    {
                        isDynamicOutPut = true;
                    }
                    break;
                default:
                {
                    if (arg.Length > _special.Length && arg.Substring(0, _special.Length) == _special)
                    {
                        special_channel = arg.Substring(_special.Length);
                    }

                    if (arg.Length > _csv_translation.Length && arg.Substring(0, _csv_translation.Length) == _csv_translation)
                    {
                        csv_translation_excel = arg.Substring(_csv_translation.Length);
                    }
                }
                break;
                }
            }

            if (!string.IsNullOrEmpty(special_channel))
            {
                Log.WriteLine(">>>>>> ���������������[{0}]", special_channel);
            }

            if (!string.IsNullOrEmpty(csv_translation_excel))
            {
                Log.WriteLine(">>>>>> csv���������ļ�[{0}]", csv_translation_excel);
            }
			
        }

        static int Main(string[] args)
        {
            ProcessCmdLine(args);

            ConvertTool convert = new ConvertTool();
            convert.BeginLoad();

            if (!fastConvert)
            {
                // Ԥ�ȼ�����Ҫ���ü����ļ�
                Log.WriteLine("��ʼǰ��У���ļ�����, ��ȴ�10����\n");
            }
			
            DirectoryInfo TheFolder = new DirectoryInfo("config");
            var xmlFiles = TheFolder.GetFiles("*.xml").Where(file => file.Name != "enums.xml");

            if (!fastConvert)
            {
                // Ԥ�ȼ�����Ҫ���ü����ļ�
                foreach (FileInfo fileInfo in xmlFiles)
                {
                    if (!convert.PreCheckLoad(fileInfo.Name))
                    {
                        GlobeError.Push("�����ļ�ʧ��:" + fileInfo.Name + "\n");
                        GlobeError.Report();
                        Console.ReadKey();
                        return -1;
                    }
                }

                Log.WriteLine("\n�����������ù�ϵ�����ϣ�ת�����ݱ�\n");
            }

            //�����ļ�
            foreach (FileInfo fileInfo in xmlFiles)
            {
                if (!convert.Convert(fileInfo.Name))
                {
                    GlobeError.Push("�����ļ�ʧ��:" + fileInfo.Name + "\n");

                    GlobeError.Report();
                    Console.ReadKey();
                    return -2;
                }
            }

            if (!i18nExtraOnly)
            {
                convert.EndLoad();
                I18N.WriteLanguageTables();
            }
            else
            {
                I18N.i18nSync();
            }

            if (GlobeError.Report())
            {
                return -3;
            }

            GlobeInfo.Report();

            return 0;
        }
    }
}
