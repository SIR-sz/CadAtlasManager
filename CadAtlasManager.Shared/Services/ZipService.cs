using System;
using System.Collections.Generic;
using System.IO;

#if CAD2012
using ICSharpCode.SharpZipLib.Zip; // R18 环境
#else
using System.IO.Compression; // R20 环境
#endif

namespace CadAtlasManager.Services
{
    public static class ZipService
    {
        // 新增：支持多文件/多文件夹打包的通用方法
        public static void CreateZipFromItems(List<Models.FileSystemItem> items, string zipFilePath)
        {
            if (File.Exists(zipFilePath)) File.Delete(zipFilePath);

#if CAD2012
            // --- R18 (SharpZipLib) 逻辑 ---
            using (ZipOutputStream s = new ZipOutputStream(File.Create(zipFilePath)))
            {
                s.SetLevel(6); // 压缩级别
                foreach (var item in items)
                {
                    if (item.Type == Models.ExplorerItemType.File)
                        AddFileToSharpZip(s, item.FullPath, item.Name);
                    else
                        AddFolderToSharpZip(s, item.FullPath, item.Name);
                }
                s.Finish();
                s.Close();
            }
#else
            // --- R20 (.NET 原生) 逻辑 ---
            using (ZipArchive archive = ZipFile.Open(zipFilePath, ZipArchiveMode.Create))
            {
                foreach (var item in items)
                {
                    if (item.Type == Models.ExplorerItemType.File)
                        archive.CreateEntryFromFile(item.FullPath, item.Name);
                    else
                        AddFolderToNativeZip(archive, item.FullPath, item.Name);
                }
            }
#endif
        }

        #region 私有辅助方法 (根据版本隔离)
#if CAD2012
        private static void AddFileToSharpZip(ZipOutputStream s, string path, string entryName) { /* SharpZipLib 具体实现 */ }
        private static void AddFolderToSharpZip(ZipOutputStream s, string path, string entryName) { /* SharpZipLib 递归 */ }
#else
        private static void AddFolderToNativeZip(ZipArchive archive, string sourceDir, string entryPrefix)
        {
            foreach (string file in Directory.GetFiles(sourceDir))
                archive.CreateEntryFromFile(file, Path.Combine(entryPrefix, Path.GetFileName(file)));
            foreach (string dir in Directory.GetDirectories(sourceDir))
                AddFolderToNativeZip(archive, dir, Path.Combine(entryPrefix, Path.GetFileName(dir)));
        }
#endif
        #endregion
    }
}