using CadAtlasManager.Models;
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

        // 【核心变更】二级字典缓存：[文件夹路径] -> [该文件夹下的 PDF 记录字典]
        // 这样可以同时管理多个 _Plot 文件夹的元数据而不会冲突
        private static Dictionary<string, Dictionary<string, string>> _folderCaches =
            new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 加载指定目录的历史记录
        /// </summary>
        public static void LoadHistory(string plotFolder)
        {
            if (string.IsNullOrEmpty(plotFolder)) return;

            var cache = new Dictionary<string, string>();
            string path = Path.Combine(plotFolder, HIDDEN_DIR, HISTORY_FILE);

            if (File.Exists(path))
            {
                try
                {
                    string[] lines = File.ReadAllLines(path);
                    foreach (var line in lines)
                    {
                        var parts = line.Split(new[] { "|||" }, StringSplitOptions.None);
                        if (parts.Length == 2) cache[parts[0]] = Decode(parts[1]);
                    }
                }
                catch { }
            }
            _folderCaches[plotFolder] = cache;
        }

        private static Dictionary<string, string> GetCache(string plotFolder)
        {
            if (!_folderCaches.ContainsKey(plotFolder)) LoadHistory(plotFolder);
            return _folderCaches[plotFolder];
        }

        // 1. 保存普通打印记录 (DWG -> PDF)
        public static void SaveRecord(string plotFolder, string pdfName, string dwgName, string fingerprint, string timestamp)
        {
            var cache = GetCache(plotFolder);
            cache[pdfName] = $"{dwgName}|{fingerprint}|{timestamp}";
            WriteToDisk(plotFolder, cache);
        }

        // [修改文件: Services/PlotMetaManager.cs]
        // 1. 修改保存合并记录的方法，改为记录快照
        public static void SaveCombinedRecord(string plotFolder, string combinedPdfName, List<FileSystemItem> sourceItems)
        {
            var cache = GetCache(plotFolder);
            // 格式：COMBINED|分项A.pdf:63851234567;分项B.pdf:63851234588
            var dataParts = sourceItems.Select(i =>
            {
                // 获取分项PDF在磁盘上的实时时间戳
                string timestamp = File.Exists(i.FullPath)
                    ? File.GetLastWriteTimeUtc(i.FullPath).Ticks.ToString()
                    : "0";
                return $"{i.Name}:{timestamp}";
            });

            cache[combinedPdfName] = COMBINED_PREFIX + string.Join(";", dataParts);
            WriteToDisk(plotFolder, cache);
        }

        // 2. 新增解析带时间戳的源文件方法
        public static Dictionary<string, string> GetCombinedSourcesWithSnapshot(string plotFolder, string pdfName)
        {
            var result = new Dictionary<string, string>();
            var cache = GetCache(plotFolder);
            if (cache.TryGetValue(pdfName, out string val) && val.StartsWith(COMBINED_PREFIX))
            {
                string raw = val.Substring(COMBINED_PREFIX.Length);
                var entries = raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var entry in entries)
                {
                    var kv = entry.Split(':');
                    if (kv.Length == 2) result[kv[0]] = kv[1];
                }
            }
            return result;
        }

        // 3. 判断是否为合并文件
        public static bool IsCombinedFile(string plotFolder, string pdfName)
        {
            var cache = GetCache(plotFolder);
            return cache.ContainsKey(pdfName) && cache[pdfName].StartsWith(COMBINED_PREFIX);
        }

        // 4. 获取合并文件的源文件列表
        public static List<string> GetCombinedSources(string plotFolder, string pdfName)
        {
            var cache = GetCache(plotFolder);
            if (cache.TryGetValue(pdfName, out string val) && val.StartsWith(COMBINED_PREFIX))
            {
                string raw = val.Substring(COMBINED_PREFIX.Length);
                return raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            }
            return new List<string>();
        }

        // 5. 获取普通文件的源 DWG 名称
        public static string GetSourceDwgName(string plotFolder, string pdfName)
        {
            var cache = GetCache(plotFolder);
            if (cache.TryGetValue(pdfName, out string val) && !val.StartsWith(COMBINED_PREFIX))
            {
                var parts = val.Split('|');
                if (parts.Length >= 1) return parts[0];
            }
            return null;
        }

        // 6. 【补全】核心校验逻辑：检查 PDF 是否过时
        public static bool CheckStatus(string plotFolder, string pdfName, string currentDwgName, string currentTimestamp, Func<string> funcToGetFingerprint)
        {
            var cache = GetCache(plotFolder);

            if (!cache.ContainsKey(pdfName)) return false;

            string savedData = cache[pdfName];
            // 如果是合并文件，这里不处理，由 AtlasView 的递归逻辑处理
            if (savedData.StartsWith(COMBINED_PREFIX)) return false;

            var parts = savedData.Split('|');
            if (parts.Length < 3) return false;

            string savedName = parts[0];
            string savedFingerprint = parts[1];
            string savedTime = parts[2];

            // 文件名校验
            if (!savedName.Equals(currentDwgName, StringComparison.OrdinalIgnoreCase)) return false;

            // 时间戳核对 (快速校验)
            if (savedTime == currentTimestamp) return true;

            // 如果时间戳变了，进一步核对 CAD 内部指纹 (Tduupdate)
            string currentFingerprint = funcToGetFingerprint();
            if (currentFingerprint == "FILE_NOT_FOUND") return false;

            if (currentFingerprint == savedFingerprint)
            {
                // 内容没变，只是文件修改时间变了 -> 更新记录以加速下次校验
                SaveRecord(plotFolder, pdfName, currentDwgName, savedFingerprint, currentTimestamp);
                return true;
            }

            return false;
        }

        private static void WriteToDisk(string plotFolder, Dictionary<string, string> cache)
        {
            try
            {
                string dir = Path.Combine(plotFolder, HIDDEN_DIR);
                if (!Directory.Exists(dir))
                {
                    DirectoryInfo di = Directory.CreateDirectory(dir);
                    di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                }

                string path = Path.Combine(dir, HISTORY_FILE);
                var lines = cache.Select(kv => $"{kv.Key}|||{Encode(kv.Value)}");
                File.WriteAllLines(path, lines);
            }
            catch { }
        }

        private static string Encode(string text) => Convert.ToBase64String(Encoding.UTF8.GetBytes(text));
        private static string Decode(string base64) => Encoding.UTF8.GetString(Convert.FromBase64String(base64));
    }
}