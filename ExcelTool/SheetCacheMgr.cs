using ExcelTool;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;

class SheetCacheMgr
{
    private static byte[] ComputeMd5(byte[] data)
    {
        return MD5.HashData(data);
    }

    private static byte[] ComputeMd5FromFile(string filename)
    {
        using FileStream stream = File.OpenRead(filename);
        return MD5.HashData(stream);
    }

    public class CacheWrite
    {
        public void AppendInt(int v)
        {
            byte[] b = BitConverter.GetBytes(v);
            b.CopyTo(data, pc);
            pc += 4;
        }

        public void AppendDouble(double v)
        {
            byte[] b = BitConverter.GetBytes(v);
            b.CopyTo(data, pc);
            pc += 4;
        }

        public void AppendString(string v)
        {
            if (v != null)
            {
                byte[] b = System.Text.Encoding.Default.GetBytes(v);
                AppendInt(b.Length);
                Array.Copy(b, 0, data, pc, b.Length);
                pc += b.Length;
            }
            else
            {
                AppendInt(0);
            }
        }

        public void SaveToFile(byte[] sourceMd5, string filename)
        {
            byte[] buff = new byte[pc + sourceMd5.Length];

            // 写入Md5
            Array.Copy(sourceMd5, 0, buff, 0, sourceMd5.Length);
            // 写入数据
            Array.Copy(data, 0, buff, sourceMd5.Length, pc);

            string dir = Path.GetDirectoryName(filename);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            if (File.Exists(filename))
            {
                FileInfo fi = new FileInfo(filename)
                {
                    Attributes = FileAttributes.Normal
                };

                File.Delete(filename);
            }
            
            File.WriteAllBytes(filename, buff);
        }

        public byte[] data = new byte[1024 * 1024 * 128]; // 最大128M
        public int pc = 0;
    }

    public static Dictionary<string, SheetCache> CachedSheets = new Dictionary<string, SheetCache>();

    const string cache_library = "cache";

    public static string getCacheFilename(string filename)
    {
        string str = string.Empty;
        byte[] bytes = ComputeMd5(System.Text.Encoding.Default.GetBytes(filename));
        for (int i = 0; i < bytes.Length; i++)
        {
            str += bytes[i].ToString("x2");
        }

        return Path.Combine(cache_library, str);
    }

    public static void AddExcelFileCache(string filename, SheetCache cache)
    {
        CachedSheets.Add(filename, cache);

        if (Program.useXlsCache)
        {
            byte[] sourceMd5 = ComputeMd5FromFile(filename);
            cache.SaveCachFile(sourceMd5, getCacheFilename(filename));
        }
    }

    public static SheetCache GetCache(string filename)
    {
        if (CachedSheets.TryGetValue(filename, out SheetCache cache))
        {
            return cache;
        }

        if (Program.useXlsCache)
        {
            string cacheFile = getCacheFilename(filename);
            if (File.Exists(cacheFile))
            {
                byte[] sourceMd5 = ComputeMd5FromFile(filename);
                byte[] cacheBytes = File.ReadAllBytes(cacheFile);
                for (int i = 0; i < sourceMd5.Length; ++i)
                {
                    if (sourceMd5[i] != cacheBytes[i])
                    {
                        return cache;
                    }
                }
                cache = new SheetCache(cacheBytes, sourceMd5.Length);
                CachedSheets.Add(filename, cache);
            }
        }

        return cache;
    }
}
