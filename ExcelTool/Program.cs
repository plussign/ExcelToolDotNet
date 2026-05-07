using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.CommandLine;

namespace ExcelTool
{
    class Program
    {
        static public bool outputLog = true;
        static public bool fastConvert = false;
        static public bool i18nExtraOnly = false;
        static public bool i18nExtraClean = false;
        static public bool isDynamicOutPut = false;
        static public bool outputCSharpAccessInterface = false;
        static public string special_channel = string.Empty;
        static public string csv_translation_excel = string.Empty;
        static public bool useXlsCache = false;
        static public bool useTestData = false;

        private class CliOptions
        {
            public bool NoLog;
            public bool Fast;
            public bool ExtraText;
            public bool CleanExtraText;
            public bool UseXlsCache;
            public bool UseTestData;
            public bool CSharp;
            public bool DynamicOutput;
            public string SpecialChannel;
            public string CsvTranslationExcel;
            public bool ShouldRun;
        }

        static string[] ExpandLegacyArgs(string[] input)
        {
            List<string> args = new List<string>();
            foreach (var s in input)
            {
                args.AddRange(s.Split('`'));
            }

            return args.ToArray();
        }

        static RootCommand BuildCommandLine(CliOptions options)
        {
            Option<bool> noLogOption = new("-nolog", "--nolog")
            {
                Description = "不输出日志"
            };
            Option<bool> fastOption = new("-fast", "--fast")
            {
                Description = "跳过前置校验，加快转换"
            };
            Option<bool> extraTextOption = new("-extra_text", "--extra-text")
            {
                Description = "只提取/同步国际化文本，并启用 fast 模式"
            };
            Option<bool> cleanExtraTextOption = new("-clean_extra_text", "--clean-extra-text")
            {
                Description = "与 -extra_text 配合，忽略已有词条并重新写入待翻译文本"
            };
            Option<bool> useXlsCacheOption = new("-use_xls_cache", "--use-xls-cache")
            {
                Description = "使用 Excel 缓存"
            };
            Option<bool> useTestDataOption = new("-use_test_data", "--use-test-data")
            {
                Description = "使用测试数据目录"
            };
            Option<bool> csharpOption = new("-csharp", "--csharp")
            {
                Description = "输出 C# 访问接口"
            };
            Option<bool> dynamicOutputOption = new("-dynamic_output", "--dynamic-output")
            {
                Description = "使用动态输出模式"
            };
            Option<string> specialOption = new("-special", "--special")
            {
                Description = "使用指定特殊渠道的输入表格，例如 -special=cn"
            };
            Option<string> csvTranslationOption = new("-csv_translation", "--csv-translation")
            {
                Description = "使用指定 CSV 词条翻译文件"
            };

            RootCommand rootCommand = new("ExcelTool 表格转换工具")
            {
                noLogOption,
                fastOption,
                extraTextOption,
                cleanExtraTextOption,
                useXlsCacheOption,
                useTestDataOption,
                csharpOption,
                dynamicOutputOption,
                specialOption,
                csvTranslationOption
            };

            rootCommand.SetAction(parseResult =>
            {
                options.NoLog = parseResult.GetValue(noLogOption);
                options.Fast = parseResult.GetValue(fastOption);
                options.ExtraText = parseResult.GetValue(extraTextOption);
                options.CleanExtraText = parseResult.GetValue(cleanExtraTextOption);
                options.UseXlsCache = parseResult.GetValue(useXlsCacheOption);
                options.UseTestData = parseResult.GetValue(useTestDataOption);
                options.CSharp = parseResult.GetValue(csharpOption);
                options.DynamicOutput = parseResult.GetValue(dynamicOutputOption);
                options.SpecialChannel = parseResult.GetValue(specialOption) ?? string.Empty;
                options.CsvTranslationExcel = parseResult.GetValue(csvTranslationOption) ?? string.Empty;
                options.ShouldRun = true;
                return 0;
            });

            return rootCommand;
        }

        static void ApplyCliOptions(CliOptions options)
        {
            outputLog = !options.NoLog;
            fastConvert = options.Fast || options.ExtraText;
            i18nExtraOnly = options.ExtraText;
            i18nExtraClean = options.CleanExtraText;
            useXlsCache = options.UseXlsCache;
            useTestData = options.UseTestData;
            outputCSharpAccessInterface = options.CSharp;
            isDynamicOutPut = options.DynamicOutput;
            special_channel = options.SpecialChannel;
            csv_translation_excel = options.CsvTranslationExcel;

            if (!string.IsNullOrEmpty(special_channel))
            {
                Log.WriteLine(">>>>>> 大区特殊输入表格[{0}]", special_channel);
            }

            if (!string.IsNullOrEmpty(csv_translation_excel))
            {
                Log.WriteLine(">>>>>> csv词条翻译文件[{0}]", csv_translation_excel);
            }
        }

        static int Main(string[] args)
        {
            CliOptions cliOptions = new CliOptions();
            RootCommand rootCommand = BuildCommandLine(cliOptions);
            int parseResult = rootCommand.Parse(ExpandLegacyArgs(args)).Invoke();
            if (parseResult != 0)
            {
                return parseResult;
            }
            if (!cliOptions.ShouldRun)
            {
                return 0;
            }

            ApplyCliOptions(cliOptions);

            ConvertTool convert = new ConvertTool();
            convert.BeginLoad();

            if (!fastConvert)
            {
                Log.WriteLine("开始前置校验文件加载，请等待10秒钟\n");
            }

            DirectoryInfo TheFolder = new DirectoryInfo("config");
            var xmlFiles = TheFolder.GetFiles("*.xml").Where(file => file.Name != "enums.xml");

            if (!fastConvert)
            {
                foreach (FileInfo fileInfo in xmlFiles)
                {
                    if (!convert.PreCheckLoad(fileInfo.Name))
                    {
                        GlobeError.Push("处理文件失败:" + fileInfo.Name + "\n");
                        GlobeError.Report();
                        Console.ReadKey();
                        return -1;
                    }
                }

                Log.WriteLine("\n表格数据引用关系检查完毕，转换数据表\n");
            }

            foreach (FileInfo fileInfo in xmlFiles)
            {
                if (!convert.Convert(fileInfo.Name))
                {
                    GlobeError.Push("处理文件失败:" + fileInfo.Name + "\n");

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
                I18N.i18nSync(i18nExtraClean);
                I18N.WriteLanguageTables();
            }

            if (GlobeError.Report())
            {
                return -3;
            }

            GlobeWarning.Report();
            GlobeInfo.Report();

            return 0;
        }
    }
}
