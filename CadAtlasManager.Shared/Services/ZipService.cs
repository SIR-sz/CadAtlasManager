using System;
using System.Collections.Generic;
using System.IO;
using CadAtlasManager.Models;

#if CAD2012
using ICSharpCode.SharpZipLib.Zip;
#else
using System.IO.Compression; 
#endif

namespace CadAtlasManager.Services
{
    public static class ZipService
    {
        public static void CreateZipFromItems(List<FileSystemItem> items, string zipFilePath)
        {
            if (File.Exists(zipFilePath)) File.Delete(zipFilePath);

#if CAD2012
            // --- R18 (SharpZipLib 0.86.0) 逻辑 ---
            using (ZipOutputStream s = new ZipOutputStream(File.Create(zipFilePath)))
            {
                s.SetLevel(6);
                byte[] buffer = new byte[4096];
                foreach (var item in items)
                {
                    if (item.Type == ExplorerItemType.File)
                        AddFileToSharpZip(s, item.FullPath, item.Name, buffer);
                    else
                        AddFolderToSharpZip(s, item.FullPath, item.Name, buffer);
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
                    if (item.Type == ExplorerItemType.File)
                        archive.CreateEntryFromFile(item.FullPath, item.Name);
                    else
                        AddFolderToNativeZip(archive, item.FullPath, item.Name);
                }
            }
#endif
        }

        #region 私有辅助方法
#if CAD2012
        private static void AddFileToSharpZip(ZipOutputStream s, string path, string entryName, byte[] buffer)
        {
            ZipEntry entry = new ZipEntry(entryName);
            entry.DateTime = DateTime.Now;
            s.PutNextEntry(entry);
            using (FileStream fs = File.OpenRead(path))
            {
                int sourceBytes;
                do
                {
                    sourceBytes = fs.Read(buffer, 0, buffer.Length);
                    s.Write(buffer, 0, sourceBytes);
                } while (sourceBytes > 0);
            }
        }

        private static void AddFolderToSharpZip(ZipOutputStream s, string path, string entryName, byte[] buffer)
        {
            foreach (string file in Directory.GetFiles(path))
                AddFileToSharpZip(s, file, Path.Combine(entryName, Path.GetFileName(file)), buffer);
            foreach (string dir in Directory.GetDirectories(path))
                AddFolderToSharpZip(s, dir, Path.Combine(entryName, Path.GetFileName(dir)), buffer);
        }
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