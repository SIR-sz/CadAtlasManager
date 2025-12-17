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

        // --- 基础 API ---
        public static string GetSmartFingerprint(string dwgPath) => "0";
        public static void InsertDwgAsBlock(string f) { }

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

        // --- 打印配置辅助 ---
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

        // --- 批量打印核心 ---
        public struct TitleBlockInfo
        {
            public string BlockName;
            public Extents3d Extents;
        }

        public static int BatchPlotByTitleBlocks(string dwgPath, string outputDir, BatchPlotConfig config)
        {
            int successCount = 0;
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
                        catch { return 0; }
                    }
                    else return 0;
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
                        if (titleBlocks.Count == 0) return 0;

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
                                    successCount++;
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
            return successCount;
        }

        private static List<TitleBlockInfo> SortTitleBlocks(List<TitleBlockInfo> list, PlotOrderType orderType)
        {
            double tolerance = 100.0;
            if (orderType == PlotOrderType.Horizontal)
            {
                return list
                    .Select(t => new { Info = t, RoundedY = Math.Round(t.Extents.MinPoint.Y / tolerance) })
                    .OrderByDescending(x => x.RoundedY)
                    .ThenBy(x => x.Info.Extents.MinPoint.X)
                    .Select(x => x.Info).ToList();
            }
            else
            {
                return list
                    .Select(t => new { Info = t, RoundedX = Math.Round(t.Extents.MinPoint.X / tolerance) })
                    .OrderBy(x => x.RoundedX)
                    .ThenByDescending(x => x.Info.Extents.MinPoint.Y)
                    .Select(x => x.Info).ToList();
            }
        }

        // --- 构建打印配置 (支持任意比例解析) ---
        private static PlotInfo BuildPlotInfo(Transaction tr, Layout layout, TitleBlockInfo tb, BatchPlotConfig config, Database targetDb)
        {
            PlotInfo info = new PlotInfo();
            info.Layout = layout.ObjectId;
            PlotSettings settings = new PlotSettings(layout.ModelType);
            settings.CopyFrom(layout);
            var psv = PlotSettingsValidator.Current;

            // 1. 打印机
            psv.SetPlotConfigurationName(settings, config.PrinterName, null);

            // 2. 纸张 (自动匹配逻辑)
            double blockW = Math.Abs(tb.Extents.MaxPoint.X - tb.Extents.MinPoint.X);
            double blockH = Math.Abs(tb.Extents.MaxPoint.Y - tb.Extents.MinPoint.Y);
            string matchedMedia = FindMatchingMedia(settings, psv, blockW, blockH);

            if (!string.IsNullOrEmpty(matchedMedia)) psv.SetCanonicalMediaName(settings, matchedMedia);
            else if (!string.IsNullOrEmpty(config.MediaName)) psv.SetCanonicalMediaName(settings, config.MediaName);

            // 3. 样式表
            if (!string.IsNullOrEmpty(config.StyleSheet))
                try { psv.SetCurrentStyleSheet(settings, config.StyleSheet); } catch { }

            // 4. 打印区域
            psv.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
            psv.SetPlotWindowArea(settings, new Extents2d(tb.Extents.MinPoint.X, tb.Extents.MinPoint.Y, tb.Extents.MaxPoint.X, tb.Extents.MaxPoint.Y));

            // 5. 自动旋转
            psv.SetPlotRotation(settings, PlotRotation.Degrees000);
            if (config.AutoRotate)
            {
                double paperW = settings.PlotPaperSize.X;
                double paperH = settings.PlotPaperSize.Y;
                bool isBlockPortrait = blockH > blockW;
                bool isPaperPortrait = paperH > paperW;
                if (isBlockPortrait != isPaperPortrait) psv.SetPlotRotation(settings, PlotRotation.Degrees090);
            }

            // 6. 比例与偏移
            if (config.ScaleType == "Fit")
            {
                psv.SetStdScaleType(settings, StdScaleType.ScaleToFit);
                psv.SetPlotCentered(settings, true);
            }
            else
            {
                // 解析自定义比例字符串 (例如 "1:100", "1:50", "0.5")
                CustomScale scale = ParseCustomScale(config.ScaleType);
                psv.SetCustomPrintScale(settings, scale);

                // 居中或偏移
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

            psv.RefreshLists(settings);
            info.OverrideSettings = settings;
            return info;
        }

        // --- 辅助方法: 解析比例字符串 ---
        private static CustomScale ParseCustomScale(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return new CustomScale(1, 1);

            try
            {
                // 处理 "1:100" 或 "1/100" 格式
                var parts = input.Split(new[] { ':', '/' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2)
                {
                    double num = double.Parse(parts[0]);
                    double den = double.Parse(parts[1]);
                    return new CustomScale(num, den);
                }
                // 处理小数 "0.5" 或 "2"
                else if (parts.Length == 1)
                {
                    double val = double.Parse(parts[0]);
                    // 如果是 0.01 (即 1:100) -> 1, 100
                    if (val < 1.0 && val > 0) return new CustomScale(1, 1.0 / val);
                    return new CustomScale(val, 1);
                }
            }
            catch { }

            return new CustomScale(1, 1); // 解析失败默认 1:1
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
                {
                    try { result.Add(new TitleBlockInfo { BlockName = effName, Extents = br.GeometricExtents }); } catch { }
                }
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
            var medias = psv.GetCanonicalMediaNameList(s);
            foreach (string mediaName in medias)
            {
                psv.SetCanonicalMediaName(s, mediaName);
                Point2d paperSize = s.PlotPaperSize;
                double pw = paperSize.X; double ph = paperSize.Y;
                if ((Math.Abs(pw - w) < tol && Math.Abs(ph - h) < tol) || (Math.Abs(pw - h) < tol && Math.Abs(ph - w) < tol))
                    return mediaName;
            }
            return null;
        }
    }
}