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

        // 格式: PDF名 -> "DwgName | Tduupdate | FileTimestamp"
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

        public static void SaveRecord(string plotFolder, string pdfName, string dwgName, string fingerprint, string timestamp)
        {
            if (!_historyCache.ContainsKey(pdfName) && _historyCache.Count == 0) LoadHistory(plotFolder);

            string data = $"{dwgName}|{fingerprint}|{timestamp}";
            _historyCache[pdfName] = data;

            WriteToDisk(plotFolder);
        }

        public static string GetSourceDwgName(string plotFolder, string pdfName)
        {
            if (_historyCache.Count == 0) LoadHistory(plotFolder);

            if (_historyCache.ContainsKey(pdfName))
            {
                var parts = _historyCache[pdfName].Split('|');
                if (parts.Length >= 1) return parts[0];
            }
            return null;
        }

        public static bool CheckStatus(string plotFolder, string pdfName, string currentDwgName, string currentTimestamp, Func<string> funcToGetFingerprint)
        {
            if (_historyCache.Count == 0) LoadHistory(plotFolder);

            if (!_historyCache.ContainsKey(pdfName)) return false;

            string savedData = _historyCache[pdfName];
            var parts = savedData.Split('|');

            if (parts.Length < 3) return false;

            string savedName = parts[0];
            string savedFingerprint = parts[1]; // 这是打印时的 Tduupdate
            string savedTime = parts[2];

            // 1. 文件名简单核对 (忽略大小写)
            if (!savedName.Equals(currentDwgName, StringComparison.OrdinalIgnoreCase)) return false;

            // 2. 如果文件系统时间戳没变，直接认为是最新的 (最快，毫秒级)
            if (savedTime == currentTimestamp) return true;

            // 3. 如果时间戳变了 (可能是复制文件，也可能是编辑了)
            // 读取当前的 Tduupdate (需要读文件头，稍慢)
            string currentFingerprint = funcToGetFingerprint();

            if (currentFingerprint == "FILE_NOT_FOUND") return false;

            // 4. 对比 Tduupdate
            if (currentFingerprint == savedFingerprint)
            {
                // Tduupdate 没变，说明只是文件被复制/移动/误触保存，但 CAD 内容没变
                // 视为最新，并自动修复记录中的时间戳，下次校验就快了
                SaveRecord(plotFolder, pdfName, currentDwgName, savedFingerprint, currentTimestamp);
                return true;
            }

            // Tduupdate 变了 -> 说明 CAD 里保存过 -> 真的过期了
            return false;
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