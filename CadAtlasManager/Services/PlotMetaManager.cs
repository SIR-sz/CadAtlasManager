using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CadAtlasManager
{
    // 专门管理图纸打印历史记录
    public static class PlotMetaManager
    {
        private const string HIDDEN_DIR = ".cadatlas";
        private const string HISTORY_FILE = "plot_history.json";

        // 记录结构：PDF文件名 -> 指纹字符串 (格式: DwgName|TdUpdateValue)
        private static Dictionary<string, string> _historyCache = new Dictionary<string, string>();

        public static void LoadHistory(string plotFolder)
        {
            _historyCache.Clear();
            string path = Path.Combine(plotFolder, HIDDEN_DIR, HISTORY_FILE);
            if (File.Exists(path))
            {
                try
                {
                    string[] lines = File.ReadAllLines(path);
                    foreach (var line in lines)
                    {
                        var parts = line.Split(new[] { "|||" }, StringSplitOptions.None);
                        if (parts.Length == 2) _historyCache[parts[0]] = Decode(parts[1]);
                    }
                }
                catch { }
            }
        }

        public static void SaveRecord(string plotFolder, string pdfName, string dwgName, string tdUpdate)
        {
            // 确保加载了当前目录的历史
            if (!_historyCache.ContainsKey(pdfName) && _historyCache.Count == 0) LoadHistory(plotFolder);

            // 指纹格式：源文件名|指纹值
            string fingerprint = $"{dwgName}|{tdUpdate}";
            _historyCache[pdfName] = fingerprint;

            WriteToDisk(plotFolder);
        }

        public static bool IsOutdated(string plotFolder, string pdfName, string currentDwgName, string currentTdUpdate)
        {
            // 如果缓存为空，尝试加载
            if (_historyCache.Count == 0) LoadHistory(plotFolder);

            if (!_historyCache.ContainsKey(pdfName)) return true; // 没有记录，视为过期(或新文件)

            string savedFingerprint = _historyCache[pdfName];
            string currentFingerprint = $"{currentDwgName}|{currentTdUpdate}";

            return savedFingerprint != currentFingerprint;
        }

        private static void WriteToDisk(string plotFolder)
        {
            string dir = Path.Combine(plotFolder, HIDDEN_DIR);
            if (!Directory.Exists(dir))
            {
                DirectoryInfo di = Directory.CreateDirectory(dir);
                di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
            }

            string path = Path.Combine(dir, HISTORY_FILE);
            var lines = _historyCache.Select(kv => $"{kv.Key}|||{Encode(kv.Value)}");
            File.WriteAllLines(path, lines);
        }

        private static string Encode(string text) => Convert.ToBase64String(Encoding.UTF8.GetBytes(text));
        private static string Decode(string base64) => Encoding.UTF8.GetString(Convert.FromBase64String(base64));
    }
}