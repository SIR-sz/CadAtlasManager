using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CadAtlasManager
{
    public static class RemarkManager
    {
        private const string HIDDEN_DIR_NAME = ".cadatlas";
        private const string META_FILE_NAME = "meta.json";

        // 缓存：Key = 文件夹路径, Value = (文件名 -> 备注内容)
        private static Dictionary<string, Dictionary<string, string>> _cache = new Dictionary<string, Dictionary<string, string>>();

        // 加载当前文件夹的备注
        public static void LoadRemarks(string folderPath)
        {
            if (_cache.ContainsKey(folderPath)) return;

            var dict = new Dictionary<string, string>();
            string metaPath = Path.Combine(folderPath, HIDDEN_DIR_NAME, META_FILE_NAME);

            if (File.Exists(metaPath))
            {
                try
                {
                    // 简易解析 JSON 格式 (避免依赖 NuGet 包)
                    string[] lines = File.ReadAllLines(metaPath);
                    foreach (var line in lines)
                    {
                        var parts = line.Split(new[] { "|||" }, StringSplitOptions.None);
                        if (parts.Length == 2)
                        {
                            dict[parts[0]] = Decode(parts[1]);
                        }
                    }
                }
                catch { }
            }
            _cache[folderPath] = dict;
        }

        // 保存/更新备注
        public static void SaveRemark(string fullPath, string remark)
        {
            string folder = Path.GetDirectoryName(fullPath);
            string fileName = Path.GetFileName(fullPath);

            if (!_cache.ContainsKey(folder)) LoadRemarks(folder);

            var dict = _cache[folder];
            if (string.IsNullOrWhiteSpace(remark))
            {
                if (dict.ContainsKey(fileName)) dict.Remove(fileName);
            }
            else
            {
                dict[fileName] = remark;
            }

            WriteToDisk(folder);
        }

        // 获取备注
        public static string GetRemark(string fullPath)
        {
            string folder = Path.GetDirectoryName(fullPath);
            string fileName = Path.GetFileName(fullPath);
            if (!_cache.ContainsKey(folder)) LoadRemarks(folder);
            return _cache[folder].TryGetValue(fileName, out string val) ? val : null;
        }

        // 处理重命名
        public static void HandleRename(string oldPath, string newPath)
        {
            string folder = Path.GetDirectoryName(oldPath);
            string oldName = Path.GetFileName(oldPath);
            string newName = Path.GetFileName(newPath);

            if (_cache.ContainsKey(folder) && _cache[folder].ContainsKey(oldName))
            {
                string content = _cache[folder][oldName];
                _cache[folder].Remove(oldName);
                _cache[folder][newName] = content;
                WriteToDisk(folder);
            }
        }

        // 处理删除
        public static void HandleDelete(string fullPath)
        {
            string folder = Path.GetDirectoryName(fullPath);
            string name = Path.GetFileName(fullPath);
            if (_cache.ContainsKey(folder) && _cache[folder].ContainsKey(name))
            {
                _cache[folder].Remove(name);
                WriteToDisk(folder);
            }
        }

        // 处理移动 (剪切粘贴/拖拽)
        public static void HandleMove(string oldPath, string newPath)
        {
            string remark = GetRemark(oldPath);
            if (!string.IsNullOrEmpty(remark))
            {
                HandleDelete(oldPath); // 旧地方删掉
                SaveRemark(newPath, remark); // 新地方加上
            }
        }

        private static void WriteToDisk(string folderPath)
        {
            if (!_cache.ContainsKey(folderPath)) return;
            var dict = _cache[folderPath];
            var hiddenDir = Path.Combine(folderPath, HIDDEN_DIR_NAME);
            var metaFile = Path.Combine(hiddenDir, META_FILE_NAME);

            try
            {
                if (dict.Count == 0)
                {
                    if (File.Exists(metaFile)) File.Delete(metaFile);
                    if (Directory.Exists(hiddenDir) && Directory.GetFileSystemEntries(hiddenDir).Length == 0)
                        Directory.Delete(hiddenDir);
                    return;
                }

                if (!Directory.Exists(hiddenDir))
                {
                    DirectoryInfo di = Directory.CreateDirectory(hiddenDir);
                    di.Attributes = FileAttributes.Directory | FileAttributes.Hidden;
                }

                // 简易保存格式：文件名|||内容 (Base64编码内容防止换行符破坏结构)
                var lines = dict.Select(kv => $"{kv.Key}|||{Encode(kv.Value)}");
                File.WriteAllLines(metaFile, lines);
            }
            catch { }
        }

        private static string Encode(string plainText) => Convert.ToBase64String(Encoding.UTF8.GetBytes(plainText));
        private static string Decode(string base64) => Encoding.UTF8.GetString(Convert.FromBase64String(base64));
    }
}