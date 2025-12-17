using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using CadAtlasManager.Models; // 引用模型
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CadAtlasManager
{
    public class CadService : IExtensionApplication
    {
        private const string FINGERPRINT_KEY = "CadAtlas_SmartFingerprint";

        // ... [Initialize, Terminate, Fingerprint 相关代码保持不变] ...
        #region 0. 插件生命周期 (请保留原有指纹代码，此处略去以节省篇幅)
        public void Initialize() { Application.DocumentManager.DocumentCreated += (s, e) => SubscribeToSaveEvent(e.Document); foreach (Document doc in Application.DocumentManager) SubscribeToSaveEvent(doc); }
        public void Terminate() { }
        private void SubscribeToSaveEvent(Document doc) { if (doc != null) doc.Database.BeginSave += Database_BeginSave; }
        private void Database_BeginSave(object sender, DatabaseIOEventArgs e) { var db = sender as Database; if (db != null) try { UpdateSmartFingerprint(db); } catch { } }
        private void UpdateSmartFingerprint(Database db) { /* ...保留原代码... */ }
        #endregion

        // ... [GetSmartFingerprint, OpenDwg, InsertDwgAsBlock 保持不变] ...
        #region 1. 基础 API (请保留)
        public static string GetSmartFingerprint(string dwgPath) { /* ...保留原代码... */ return "0"; }
        public static void OpenDwg(string sourcePath, string mode, string targetCopyPath = null)
        {
            if (!File.Exists(sourcePath))
            {
                // 【新增】如果文件找不到，弹出提示，而不是没反应
                System.Windows.MessageBox.Show($"文件不存在或路径错误：\n{sourcePath}", "打开失败");
                return;
            }
        }
        public static void InsertDwgAsBlock(string f) { /* ...保留原代码... */ }
        #endregion

        // =================================================================
        // 【新增】辅助方法：获取系统打印配置列表 (供 UI 调用)
        // =================================================================
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

        // =================================================================
        // 【重构】批量打印核心 (使用 BatchPlotConfig)
        // =================================================================

        public struct TitleBlockInfo
        {
            public string BlockName;
            public Extents3d Extents;
        }

        // 修改：参数改为 BatchPlotConfig
        public static int BatchPlotByTitleBlocks(string dwgPath, string outputDir, BatchPlotConfig config)
        {
            int successCount = 0;
            Document doc = null;
            bool isOpenedByUs = false;

            try
            {
                // A. 安全打开文档
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

                if (Application.DocumentManager.MdiActiveDocument != doc) Application.DocumentManager.MdiActiveDocument = doc;

                Database db = doc.Database;

                // B. 解析图框名列表
                List<string> blockNames = config.TitleBlockNames
                    .Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(s => s.Trim()).ToList();

                using (doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // 1. 定位空间 & 扫描图框
                    ObjectId targetSpaceId = db.CurrentSpaceId;
                    var targetLayout = GetLayoutFromBtrId(tr, db, targetSpaceId);
                    if (targetLayout == null) return 0;

                    // 切换环境
                    if (LayoutManager.Current.CurrentLayout != targetLayout.LayoutName)
                        LayoutManager.Current.CurrentLayout = targetLayout.LayoutName;

                    var titleBlocks = ScanTitleBlocks(tr, targetSpaceId, blockNames);
                    if (titleBlocks.Count == 0) return 0;

                    // 2. 执行打印
                    if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                    {
                        using (PlotEngine engine = PlotFactory.CreatePublishEngine())
                        {
                            engine.BeginPlot(null, null);

                            for (int i = 0; i < titleBlocks.Count; i++)
                            {
                                var tb = titleBlocks[i];
                                string fileName = Path.GetFileNameWithoutExtension(dwgPath);
                                if (titleBlocks.Count > 1) fileName += $"_{i + 1}";
                                string fullPdfPath = Path.Combine(outputDir, fileName + ".pdf");

                                // 3. 构建配置 (传入 config)
                                PlotInfo plotInfo = BuildPlotInfo(tr, targetLayout, tb, config, db);

                                PlotInfoValidator validator = new PlotInfoValidator();
                                validator.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                                validator.Validate(plotInfo);

                                engine.BeginDocument(plotInfo, doc.Name, null, 1, true, fullPdfPath);
                                PlotPageInfo pageInfo = new PlotPageInfo();
                                engine.BeginPage(pageInfo, plotInfo, true, null);
                                engine.BeginGenerateGraphics(null);
                                engine.EndGenerateGraphics(null);
                                engine.EndPage(null);
                                engine.EndDocument(null);

                                successCount++;
                            }
                            engine.EndPlot(null);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show($"打印失败 {Path.GetFileName(dwgPath)}: {ex.Message}");
            }
            finally
            {
                if (isOpenedByUs && doc != null) doc.CloseAndDiscard();
            }

            return successCount;
        }

        // 修改：构建配置的方法
        private static PlotInfo BuildPlotInfo(Transaction tr, Layout layout, TitleBlockInfo tb, BatchPlotConfig config, Database targetDb)
        {
            PlotInfo info = new PlotInfo();
            info.Layout = layout.ObjectId;

            PlotSettings settings = new PlotSettings(layout.ModelType);
            settings.CopyFrom(layout); // 保留线宽开关等基础设置

            var psv = PlotSettingsValidator.Current;

            // 1. 设置打印机
            try { psv.SetPlotConfigurationName(settings, config.PrinterName, null); }
            catch { psv.SetPlotConfigurationName(settings, "None_Device", null); }

            // 2. 匹配纸张
            double w = Math.Abs(tb.Extents.MaxPoint.X - tb.Extents.MinPoint.X);
            double h = Math.Abs(tb.Extents.MaxPoint.Y - tb.Extents.MinPoint.Y);
            string matched = FindMatchingMedia(settings, psv, w, h);

            if (!string.IsNullOrEmpty(matched)) psv.SetCanonicalMediaName(settings, matched);
            else if (!string.IsNullOrEmpty(config.MediaName)) psv.SetCanonicalMediaName(settings, config.MediaName);

            // 3. 样式表 (安全检查)
            if (!string.IsNullOrEmpty(config.StyleSheet))
            {
                try
                {
                    bool isCtbFile = config.StyleSheet.EndsWith(".ctb", StringComparison.OrdinalIgnoreCase);
                    bool isDrawingCtb = targetDb.PlotStyleMode;
                    if (isCtbFile == isDrawingCtb) psv.SetCurrentStyleSheet(settings, config.StyleSheet);
                }
                catch { }
            }

            // 4. 旋转与窗口
            if (config.AutoRotate)
            {
                // 简单逻辑：如果纸张匹配到了，通常驱动会自动处理旋转。
                // 如果需要强制逻辑，可以在此扩展。目前保持 UseLast 或 0 度
                // psv.SetPlotRotation(settings, PlotRotation.Degrees000); 
            }

            psv.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
            psv.SetPlotWindowArea(settings, new Extents2d(tb.Extents.MinPoint.X, tb.Extents.MinPoint.Y, tb.Extents.MaxPoint.X, tb.Extents.MaxPoint.Y));

            if (config.UseStandardScale)
            {
                psv.SetStdScaleType(settings, StdScaleType.ScaleToFit);
                psv.SetPlotCentered(settings, true);
            }

            psv.RefreshLists(settings);
            info.OverrideSettings = settings;
            return info;
        }

        // ... [辅助方法 FindMatchingMedia, ScanTitleBlocks, GetEffectiveName 等保持不变，参考上一次回答] ...
        // 为确保完整性，请保留上一次回答中的 FindMatchingMedia 等辅助方法
        private static string FindMatchingMedia(PlotSettings s, PlotSettingsValidator psv, double w, double h) { /*...参考上文...*/ return null; }
        private static List<TitleBlockInfo> ScanTitleBlocks(Transaction tr, ObjectId id, List<string> names) { /*...参考上文...*/ return new List<TitleBlockInfo>(); }
        private static Layout GetLayoutFromBtrId(Transaction tr, Database db, ObjectId id) { /*...参考上文...*/ return null; }
        private static string GetEffectiveName(BlockReference br, Transaction tr) { /*...参考上文...*/ return null; }
        private static bool IsSize(double l1, double s1, double l2, double s2, double t) { return Math.Abs(l1 - l2) < t && Math.Abs(s1 - s2) < t; }
    }
}