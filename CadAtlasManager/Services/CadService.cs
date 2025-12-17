using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using CadAtlasManager.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CadAtlasManager
{
    public class CadService : IExtensionApplication
    {
        public void Initialize() { }
        public void Terminate() { }

        // 核心：获取 CAD 内部的最后保存时间 (Tduupdate) 作为指纹
        public static string GetContentFingerprint(string dwgPath)
        {
            // A. 如果文件已打开，从内存读
            foreach (Document doc in Application.DocumentManager)
            {
                if (doc.Name.Equals(dwgPath, StringComparison.OrdinalIgnoreCase))
                {
                    // 使用 Tduupdate (Universal Time)，解决属性名报错问题
                    return doc.Database.Tduupdate.ToString();
                }
            }

            // B. 如果文件未打开，使用侧数据库静默读取 (ReadDwgFile)
            if (File.Exists(dwgPath))
            {
                try
                {
                    using (Database db = new Database(false, true))
                    {
                        db.ReadDwgFile(dwgPath, FileShare.Read, true, "");
                        return db.Tduupdate.ToString();
                    }
                }
                catch
                {
                    // 降级方案：文件系统时间
                    return File.GetLastWriteTimeUtc(dwgPath).Ticks.ToString();
                }
            }
            return "FILE_NOT_FOUND";
        }

        public static string GetFileTimestamp(string dwgPath)
        {
            if (!File.Exists(dwgPath)) return "0";
            return File.GetLastWriteTimeUtc(dwgPath).Ticks.ToString();
        }

        public static void InsertDwgAsBlock(string f)
        {
            if (!File.Exists(f)) return;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            doc.Window.Focus();
            string cmd = $"-INSERT \"{f}\" ";
            doc.SendStringToExecute(cmd, true, false, false);
        }

        public static void OpenDwg(string sourcePath, string mode, string targetCopyPath = null)
        {
            if (!File.Exists(sourcePath)) { System.Windows.MessageBox.Show($"文件不存在：\n{sourcePath}"); return; }
            DocumentCollection acDocMgr = Application.DocumentManager;
            string finalOpenPath = sourcePath;
            bool isReadOnly = (mode == "Read");

            if (mode == "Copy" && !string.IsNullOrEmpty(targetCopyPath))
            {
                try { File.Copy(sourcePath, targetCopyPath, true); finalOpenPath = targetCopyPath; isReadOnly = false; }
                catch (System.Exception ex) { System.Windows.MessageBox.Show($"复制失败: {ex.Message}"); return; }
            }

            foreach (Document doc in acDocMgr)
            {
                if (doc.Name.Equals(finalOpenPath, StringComparison.OrdinalIgnoreCase)) { acDocMgr.MdiActiveDocument = doc; return; }
            }
            try { acDocMgr.Open(finalOpenPath, isReadOnly); }
            catch (System.Exception ex) { System.Windows.MessageBox.Show($"CAD 无法打开文件:\n{ex.Message}"); }
        }

        // --- 打印辅助方法 ---
        public static List<string> GetPlotters()
        {
            try { return PlotSettingsValidator.Current.GetPlotDeviceList().Cast<string>().ToList(); }
            catch { return new List<string>(); }
        }
        public static List<string> GetMediaList(string deviceName)
        {
            try
            {
                using (PlotSettings temp = new PlotSettings(true))
                {
                    PlotSettingsValidator.Current.SetPlotConfigurationName(temp, deviceName, null);
                    return PlotSettingsValidator.Current.GetCanonicalMediaNameList(temp).Cast<string>().ToList();
                }
            }
            catch { return new List<string>(); }
        }
        public static List<string> GetStyleSheets()
        {
            try { return PlotSettingsValidator.Current.GetPlotStyleSheetList().Cast<string>().ToList(); }
            catch { return new List<string>(); }
        }

        public struct TitleBlockInfo
        {
            public string BlockName;
            public Extents3d Extents;
        }

        // --- 批量打印核心 ---
        // 修改：返回 List<string> 包含所有生成的 PDF 文件名
        public static List<string> BatchPlotByTitleBlocks(string dwgPath, string outputDir, BatchPlotConfig config)
        {
            List<string> generatedFiles = new List<string>();
            Document doc = null;
            bool isOpenedByUs = false;

            try
            {
                foreach (Document d in Application.DocumentManager)
                {
                    if (d.Name.Equals(dwgPath, StringComparison.OrdinalIgnoreCase)) { doc = d; break; }
                }
                if (doc == null)
                {
                    if (File.Exists(dwgPath))
                    {
                        try { doc = Application.DocumentManager.Open(dwgPath, false); isOpenedByUs = true; }
                        catch { return generatedFiles; }
                    }
                    else return generatedFiles;
                }

                using (doc.LockDocument())
                {
                    if (Application.DocumentManager.MdiActiveDocument != doc)
                        Application.DocumentManager.MdiActiveDocument = doc;

                    Database db = doc.Database;
                    List<string> blockNames = config.TitleBlockNames.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        ObjectId targetSpaceId = db.CurrentSpaceId;
                        var targetLayout = GetLayoutFromBtrId(tr, db, targetSpaceId);

                        var titleBlocks = ScanTitleBlocks(tr, targetSpaceId, blockNames);
                        if (titleBlocks.Count == 0) return generatedFiles;

                        titleBlocks = SortTitleBlocks(titleBlocks, config.OrderType);

                        for (int i = 0; i < titleBlocks.Count; i++)
                        {
                            var tb = titleBlocks[i];
                            string fileName = Path.GetFileNameWithoutExtension(dwgPath);
                            if (titleBlocks.Count > 1) fileName += $"_{i + 1}";
                            string fullPdfPath = Path.Combine(outputDir, fileName + ".pdf");

                            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                            {
                                using (PlotEngine engine = PlotFactory.CreatePublishEngine())
                                {
                                    PlotInfo plotInfo = BuildPlotInfo(tr, targetLayout, tb, config, db);
                                    PlotInfoValidator validator = new PlotInfoValidator();
                                    validator.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                                    validator.Validate(plotInfo);

                                    engine.BeginPlot(null, null);
                                    engine.BeginDocument(plotInfo, doc.Name, null, 1, true, fullPdfPath);
                                    engine.BeginPage(new PlotPageInfo(), plotInfo, true, null);
                                    engine.BeginGenerateGraphics(null);
                                    engine.EndGenerateGraphics(null);
                                    engine.EndPage(null);
                                    engine.EndDocument(null);
                                    engine.EndPlot(null);

                                    generatedFiles.Add(fileName + ".pdf"); // 记录生成的文件
                                }
                            }
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show($"打印出错: {ex.Message}");
            }
            finally
            {
                if (isOpenedByUs && doc != null) doc.CloseAndDiscard();
            }
            return generatedFiles;
        }

        // ... (SortTitleBlocks, BuildPlotInfo, ParseCustomScale, ScanTitleBlocks 等辅助方法与之前保持一致) ...
        private static List<TitleBlockInfo> SortTitleBlocks(List<TitleBlockInfo> list, PlotOrderType orderType)
        {
            double tolerance = 100.0;
            if (orderType == PlotOrderType.Horizontal)
            {
                return list.Select(t => new { Info = t, RoundedY = Math.Round(t.Extents.MinPoint.Y / tolerance) })
                    .OrderByDescending(x => x.RoundedY).ThenBy(x => x.Info.Extents.MinPoint.X).Select(x => x.Info).ToList();
            }
            else
            {
                return list.Select(t => new { Info = t, RoundedX = Math.Round(t.Extents.MinPoint.X / tolerance) })
                    .OrderBy(x => x.RoundedX).ThenByDescending(x => x.Info.Extents.MinPoint.Y).Select(x => x.Info).ToList();
            }
        }
        private static PlotInfo BuildPlotInfo(Transaction tr, Layout layout, TitleBlockInfo tb, BatchPlotConfig config, Database targetDb)
        {
            PlotInfo info = new PlotInfo();
            info.Layout = layout.ObjectId;

            // 1. 初始化
            PlotSettings settings = new PlotSettings(layout.ModelType);
            settings.CopyFrom(layout);
            var psv = PlotSettingsValidator.Current;

            // 2. 设置打印机
            psv.SetPlotConfigurationName(settings, config.PrinterName, null);

            // 3. 计算图框宽高
            double blockW = Math.Abs(tb.Extents.MaxPoint.X - tb.Extents.MinPoint.X);
            double blockH = Math.Abs(tb.Extents.MaxPoint.Y - tb.Extents.MinPoint.Y);

            // 4. 纸张匹配 (调用上面的增强版方法)
            string matchedMedia = FindMatchingMedia(settings, psv, blockW, blockH);

            if (!string.IsNullOrEmpty(matchedMedia))
            {
                // 自动匹配成功
                psv.SetCanonicalMediaName(settings, matchedMedia);
            }
            else
            {
                // 匹配失败，尝试使用用户指定的纸张兜底
                if (!string.IsNullOrEmpty(config.MediaName))
                {
                    try { psv.SetCanonicalMediaName(settings, config.MediaName); } catch { }
                }
            }

            // 5. 样式表 (CTB)
            if (!string.IsNullOrEmpty(config.StyleSheet))
            {
                try { psv.SetCurrentStyleSheet(settings, config.StyleSheet); } catch { }
            }

            // 6. 打印区域 (Window)
            psv.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
            psv.SetPlotWindowArea(settings, new Extents2d(tb.Extents.MinPoint.X, tb.Extents.MinPoint.Y, tb.Extents.MaxPoint.X, tb.Extents.MaxPoint.Y));

            // 7. 旋转
            psv.SetPlotRotation(settings, PlotRotation.Degrees000);
            if (config.AutoRotate)
            {
                double paperW = settings.PlotPaperSize.X;
                double paperH = settings.PlotPaperSize.Y;
                if ((blockH > blockW) != (paperH > paperW))
                {
                    psv.SetPlotRotation(settings, PlotRotation.Degrees090);
                }
            }

            // 8. 比例与居中
            if (config.ScaleType == "Fit")
            {
                psv.SetStdScaleType(settings, StdScaleType.ScaleToFit);
                psv.SetPlotCentered(settings, true);
            }
            else
            {
                CustomScale scale = ParseCustomScale(config.ScaleType);
                psv.SetCustomPrintScale(settings, scale);

                if (config.PlotCentered)
                {
                    psv.SetPlotCentered(settings, true);
                }
                else
                {
                    psv.SetPlotCentered(settings, false);
                    psv.SetPlotOrigin(settings, new Point2d(config.OffsetX, config.OffsetY));
                }
            }

            // 9. 最终应用
            psv.RefreshLists(settings);
            info.OverrideSettings = settings;

            return info;
        }
        private static CustomScale ParseCustomScale(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return new CustomScale(1, 1);
            try
            {
                var parts = input.Split(new[] { ':', '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2) return new CustomScale(double.Parse(parts[0]), double.Parse(parts[1]));
                else if (parts.Length == 1) { double val = double.Parse(parts[0]); return val < 1.0 && val > 0 ? new CustomScale(1, 1.0 / val) : new CustomScale(val, 1); }
            }
            catch { }
            return new CustomScale(1, 1);
        }
        private static List<TitleBlockInfo> ScanTitleBlocks(Transaction tr, ObjectId spaceId, List<string> targetNames)
        {
            var result = new List<TitleBlockInfo>();
            var btr = tr.GetObject(spaceId, OpenMode.ForRead) as BlockTableRecord;
            if (btr == null) return result;
            foreach (ObjectId id in btr)
            {
                if (id.ObjectClass.DxfName != "INSERT") continue;
                BlockReference br = tr.GetObject(id, OpenMode.ForRead) as BlockReference;
                if (br == null) continue;
                string effName = GetEffectiveName(br);
                if (targetNames.Any(n => n.Equals(effName, StringComparison.OrdinalIgnoreCase)))
                    try { result.Add(new TitleBlockInfo { BlockName = effName, Extents = br.GeometricExtents }); } catch { }
            }
            return result;
        }
        private static string GetEffectiveName(BlockReference br)
        {
            if (br.IsDynamicBlock)
            {
                using (var tr = br.Database.TransactionManager.StartTransaction())
                {
                    var btr = tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    return btr?.Name ?? br.Name;
                }
            }
            return br.Name;
        }
        private static Layout GetLayoutFromBtrId(Transaction tr, Database db, ObjectId btrId)
        {
            var btr = tr.GetObject(btrId, OpenMode.ForRead) as BlockTableRecord;
            if (btr == null || !btr.IsLayout) return null;
            return tr.GetObject(btr.LayoutId, OpenMode.ForRead) as Layout;
        }
        private static string FindMatchingMedia(PlotSettings s, PlotSettingsValidator psv, double w, double h)
        {
            double tol = 2.0;
            // 1. 备份当前纸张 (如果它是有效的)
            string originalMedia = s.CanonicalMediaName;

            // 获取打印机支持的所有纸张列表
            var medias = psv.GetCanonicalMediaNameList(s);
            if (medias == null || medias.Count == 0) return null;

            // 2. 遍历查找匹配尺寸
            foreach (string mediaName in medias)
            {
                psv.SetCanonicalMediaName(s, mediaName); // 临时切换纸张以获取尺寸
                Point2d paperSize = s.PlotPaperSize;
                double pw = paperSize.X;
                double ph = paperSize.Y;

                // 检查尺寸 (横向或纵向)
                if ((Math.Abs(pw - w) < tol && Math.Abs(ph - h) < tol) ||
                    (Math.Abs(pw - h) < tol && Math.Abs(ph - w) < tol))
                {
                    return mediaName; // 找到了！此时 s 已设置为正确纸张
                }
            }

            // 3. 【关键防崩逻辑】如果没找到匹配的...
            // 此时 s 可能停留在列表的最后一张纸（往往是 UserDefined 或 Custom），这会导致 eInvalidInput 报错。
            // 必须把它“救”回来！

            bool restored = false;
            // A. 尝试还原回原来的纸张 (如果原来的在列表里)
            if (!string.IsNullOrEmpty(originalMedia) && medias.Contains(originalMedia))
            {
                try { psv.SetCanonicalMediaName(s, originalMedia); restored = true; } catch { }
            }

            // B. 如果还原失败（比如切了打印机，旧纸张不支持），强制设置为列表的第一张纸 (如 A4)
            // 这样虽然尺寸不对，但至少是一个“有效”的纸张，能保证 Validator 不报错。
            if (!restored && medias.Count > 0)
            {
                try { psv.SetCanonicalMediaName(s, medias[0]); } catch { }
            }

            return null; // 返回 null 表示没找到完美匹配的
        }
    }
}