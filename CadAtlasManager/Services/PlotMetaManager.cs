using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CadAtlasManager
{
    public static class PlotMetaManager
    {
        private const string HIDDEN_DIR = ".cadatlas";
        private const string HISTORY_FILE = "plot_history.json";
        private const string COMBINED_PREFIX = "COMBINED|";

        // 缓存字典
        // 普通记录: PDF名 -> "DwgName | Tduupdate | FileTimestamp"
        // 合并记录: PDF名 -> "COMBINED|Source1.pdf;Source2.pdf;..."
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

        // 1. 保存普通打印记录 (DWG -> PDF)
        public static void SaveRecord(string plotFolder, string pdfName, string dwgName, string fingerprint, string timestamp)
        {
            EnsureLoaded(plotFolder);
            string data = $"{dwgName}|{fingerprint}|{timestamp}";
            _historyCache[pdfName] = data;
            WriteToDisk(plotFolder);
        }

        // 2. 【新增】保存合并记录 (List<PDF> -> CombinedPDF)
        public static void SaveCombinedRecord(string plotFolder, string combinedPdfName, List<string> sourcePdfNames)
        {
            EnsureLoaded(plotFolder);
            // 格式: COMBINED|Source1.pdf;Source2.pdf
            string data = COMBINED_PREFIX + string.Join(";", sourcePdfNames);
            _historyCache[combinedPdfName] = data;
            WriteToDisk(plotFolder);
        }

        // 3. 【新增】判断是否为合并文件
        public static bool IsCombinedFile(string plotFolder, string pdfName)
        {
            EnsureLoaded(plotFolder);
            return _historyCache.ContainsKey(pdfName) && _historyCache[pdfName].StartsWith(COMBINED_PREFIX);
        }

        // 4. 【新增】获取合并文件的源文件列表
        public static List<string> GetCombinedSources(string plotFolder, string pdfName)
        {
            EnsureLoaded(plotFolder);
            if (_historyCache.TryGetValue(pdfName, out string val))
            {
                if (val.StartsWith(COMBINED_PREFIX))
                {
                    string raw = val.Substring(COMBINED_PREFIX.Length);
                    return raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                }
            }
            return new List<string>();
        }

        // 5. 获取普通文件的源 DWG 名称
        public static string GetSourceDwgName(string plotFolder, string pdfName)
        {
            EnsureLoaded(plotFolder);
            if (_historyCache.ContainsKey(pdfName))
            {
                string val = _historyCache[pdfName];
                if (!val.StartsWith(COMBINED_PREFIX))
                {
                    var parts = val.Split('|');
                    if (parts.Length >= 1) return parts[0];
                }
            }
            return null;
        }

        // 6. 核心校验逻辑 (针对单张 DWG 导出的 PDF)
        public static bool CheckStatus(string plotFolder, string pdfName, string currentDwgName, string currentTimestamp, Func<string> funcToGetFingerprint)
        {
            EnsureLoaded(plotFolder);

            if (!_historyCache.ContainsKey(pdfName)) return false;

            string savedData = _historyCache[pdfName];
            // 如果是合并文件，这里不处理，返回 false
            if (savedData.StartsWith(COMBINED_PREFIX)) return false;

            var parts = savedData.Split('|');
            if (parts.Length < 3) return false;

            string savedName = parts[0];
            string savedFingerprint = parts[1];
            string savedTime = parts[2];

            // 1. 文件名核对
            if (!savedName.Equals(currentDwgName, StringComparison.OrdinalIgnoreCase)) return false;

            // 2. 时间戳核对 (最快)
            if (savedTime == currentTimestamp) return true;

            // 3. 读取 Tduupdate 核对 (较慢但准确)
            string currentFingerprint = funcToGetFingerprint();
            if (currentFingerprint == "FILE_NOT_FOUND") return false;

            if (currentFingerprint == savedFingerprint)
            {
                // 内容没变，只是时间变了 -> 更新记录以加速下次校验
                SaveRecord(plotFolder, pdfName, currentDwgName, savedFingerprint, currentTimestamp);
                return true;
            }

            return false;
        }

        private static void EnsureLoaded(string plotFolder)
        {
            if (_historyCache.Count == 0) LoadHistory(plotFolder);
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