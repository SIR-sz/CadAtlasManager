using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CadAtlasManager
{
    /// <summary>
    /// 核心 CAD 服务类 - 增加“自动识别图框大小并匹配纸张”功能
    /// </summary>
    public class CadService : IExtensionApplication
    {
        private const string FINGERPRINT_KEY = "CadAtlas_SmartFingerprint";

        #region 0. 插件生命周期管理
        public void Initialize()
        {
            Application.DocumentManager.DocumentCreated += (s, e) => SubscribeToSaveEvent(e.Document);
            foreach (Document doc in Application.DocumentManager) SubscribeToSaveEvent(doc);
        }

        public void Terminate() { }

        private void SubscribeToSaveEvent(Document doc)
        {
            if (doc != null) doc.Database.BeginSave += Database_BeginSave;
        }

        private void Database_BeginSave(object sender, DatabaseIOEventArgs e)
        {
            var db = sender as Database;
            if (db == null) return;
            try { UpdateSmartFingerprint(db); } catch { }
        }

        private void UpdateSmartFingerprint(Database db)
        {
            try
            {
                DatabaseSummaryInfoBuilder info = new DatabaseSummaryInfoBuilder(db.SummaryInfo);
                if (info.CustomPropertyTable.Contains(FINGERPRINT_KEY))
                    info.CustomPropertyTable[FINGERPRINT_KEY] = DateTime.Now.ToString("o");
                else
                    info.CustomPropertyTable.Add(FINGERPRINT_KEY, DateTime.Now.ToString("o"));
                db.SummaryInfo = info.ToDatabaseSummaryInfo();
            }
            catch { }
        }
        #endregion

        #region 1. 对外基础 API

        public static string GetSmartFingerprint(string dwgPath)
        {
            if (!File.Exists(dwgPath)) return "0";
            try
            {
                using (Database db = new Database(false, true))
                {
                    db.ReadDwgFile(dwgPath, FileShare.Read, false, null);
                    DatabaseSummaryInfo info = db.SummaryInfo;
                    System.Collections.IDictionaryEnumerator iter = info.CustomProperties;
                    while (iter.MoveNext())
                    {
                        if (iter.Key.ToString() == FINGERPRINT_KEY) return iter.Value.ToString();
                    }
                    return db.Tdupdate.ToString("o");
                }
            }
            catch { return "0"; }
        }

        public static void OpenDwg(string sourcePath, string mode, string targetCopyPath = null)
        {
            if (!File.Exists(sourcePath)) return;
            DocumentCollection acDocMgr = Application.DocumentManager;
            string finalPath = sourcePath;
            bool readOnly = (mode == "Read");

            if (mode == "Copy" && !string.IsNullOrEmpty(targetCopyPath))
            {
                File.Copy(sourcePath, targetCopyPath, true);
                finalPath = targetCopyPath;
                readOnly = false;
            }

            foreach (Document doc in acDocMgr)
            {
                if (doc.Name.Equals(finalPath, StringComparison.OrdinalIgnoreCase))
                {
                    acDocMgr.MdiActiveDocument = doc;
                    return;
                }
            }
            acDocMgr.Open(finalPath, readOnly);
        }

        public static void InsertDwgAsBlock(string filePath)
        {
            if (!File.Exists(filePath)) return;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            try
            {
                using (doc.LockDocument())
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    doc.Editor.Command("_.INSERT", filePath, "0,0", "1", "1", "0");
                    tr.Commit();
                }
            }
            catch { }
        }

        public static PlotSettings GetTemplatePlotSettings()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return null;

            using (doc.LockDocument())
            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                var layoutMgr = LayoutManager.Current;
                var layoutId = layoutMgr.GetLayoutId(layoutMgr.CurrentLayout);
                var layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);

                var settings = new PlotSettings(layout.ModelType);
                settings.CopyFrom(layout);
                tr.Commit();
                return settings;
            }
        }
        #endregion

        #region 2. 批量打印核心

        public struct TitleBlockInfo
        {
            public string BlockName;
            public Extents3d Extents;
        }

        public static int BatchPlotByTitleBlocks(string dwgPath, string outputDir, List<string> blockNames, PlotSettings templateSettings)
        {
            int successCount = 0;
            Document doc = null;
            bool isOpenedByUs = false;

            try
            {
                // A. 获取或打开文档
                foreach (Document d in Application.DocumentManager)
                {
                    if (d.Name.Equals(dwgPath, StringComparison.OrdinalIgnoreCase))
                    {
                        doc = d;
                        break;
                    }
                }

                if (doc == null)
                {
                    if (File.Exists(dwgPath))
                    {
                        try
                        {
                            doc = Application.DocumentManager.Open(dwgPath, false);
                            isOpenedByUs = true;
                        }
                        catch (System.Exception openEx)
                        {
                            System.Windows.MessageBox.Show($"无法打开文件: {Path.GetFileName(dwgPath)}\n{openEx.Message}");
                            return 0;
                        }
                    }
                    else return 0;
                }

                if (Application.DocumentManager.MdiActiveDocument != doc)
                {
                    Application.DocumentManager.MdiActiveDocument = doc;
                }

                Database db = doc.Database;

                // B. 锁定并执行
                using (doc.LockDocument())
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    ObjectId targetSpaceId = db.CurrentSpaceId;
                    var targetLayout = GetLayoutFromBtrId(tr, db, targetSpaceId);

                    if (targetLayout == null)
                    {
                        System.Diagnostics.Debug.WriteLine("无法解析当前布局，跳过文件。");
                        return 0;
                    }

                    if (LayoutManager.Current.CurrentLayout != targetLayout.LayoutName)
                    {
                        LayoutManager.Current.CurrentLayout = targetLayout.LayoutName;
                    }

                    var titleBlocks = ScanTitleBlocks(tr, targetSpaceId, blockNames);
                    if (titleBlocks.Count == 0) return 0;

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

                                // 5. 构建配置 (这里增加了自动纸张匹配逻辑)
                                PlotInfo plotInfo = BuildPlotInfo(tr, targetLayout, tb, templateSettings, db);

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
                System.Windows.MessageBox.Show($"文件 {Path.GetFileName(dwgPath)} 打印失败:\n{ex.Message}");
            }
            finally
            {
                if (isOpenedByUs && doc != null)
                {
                    doc.CloseAndDiscard();
                }
            }

            return successCount;
        }

        // --- 辅助方法 ---

        private static List<TitleBlockInfo> ScanTitleBlocks(Transaction tr, ObjectId spaceId, List<string> targetNames)
        {
            var list = new List<TitleBlockInfo>();
            var btr = (BlockTableRecord)tr.GetObject(spaceId, OpenMode.ForRead);

            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent is BlockReference br)
                {
                    string name = GetEffectiveName(br, tr);
                    if (!string.IsNullOrEmpty(name) &&
                        targetNames.Any(n => n.Equals(name, StringComparison.OrdinalIgnoreCase)))
                    {
                        list.Add(new TitleBlockInfo { BlockName = name, Extents = br.GeometricExtents });
                    }
                }
            }
            return list.OrderByDescending(x => x.Extents.MaxPoint.Y)
                       .ThenBy(x => x.Extents.MinPoint.X)
                       .ToList();
        }

        // 【核心修改】自动匹配纸张大小逻辑
        private static PlotInfo BuildPlotInfo(Transaction tr, Layout layout, TitleBlockInfo tb, PlotSettings template, Database targetDb)
        {
            PlotInfo info = new PlotInfo();
            info.Layout = layout.ObjectId;

            PlotSettings settings = new PlotSettings(layout.ModelType);
            settings.CopyFrom(layout);

            var psv = PlotSettingsValidator.Current;

            // A. 设置打印机 (必须先设置打印机，才能查询纸张)
            psv.SetPlotConfigurationName(settings, template.PlotConfigurationName, null); // 先不设纸张，设为null

            // B. 智能匹配纸张 (Auto-Match Paper Size)
            // 1. 计算图框尺寸 (假设单位是毫米，容差 10mm)
            double width = Math.Abs(tb.Extents.MaxPoint.X - tb.Extents.MinPoint.X);
            double height = Math.Abs(tb.Extents.MaxPoint.Y - tb.Extents.MinPoint.Y);

            // 2. 尝试从打印机驱动中寻找匹配的纸张
            string matchedMedia = FindMatchingMedia(settings, psv, width, height);

            if (!string.IsNullOrEmpty(matchedMedia))
            {
                // 找到了匹配的纸张 (例如 A3)
                psv.SetCanonicalMediaName(settings, matchedMedia);
            }
            else
            {
                // 没找到，回退到模板的默认纸张
                psv.SetCanonicalMediaName(settings, template.CanonicalMediaName);
            }

            // C. 样式表
            if (!string.IsNullOrEmpty(template.CurrentStyleSheet))
            {
                try
                {
                    bool isCtbFile = template.CurrentStyleSheet.EndsWith(".ctb", StringComparison.OrdinalIgnoreCase);
                    bool isDrawingCtb = targetDb.PlotStyleMode;
                    if (isCtbFile == isDrawingCtb)
                    {
                        psv.SetCurrentStyleSheet(settings, template.CurrentStyleSheet);
                    }
                }
                catch { }
            }

            // D. 旋转 (根据图框长宽比自动决定横纵向)
            // 如果图框宽 > 高，通常是横向 (Landscape)
            // 但如果用户模板强制了旋转，这里可能需要权衡。
            // 策略：如果找到了匹配纸张，使用纸张的自然方向；否则沿用模板。
            // 这里我们保持模板的旋转设置，但在匹配纸张时已经考虑了长宽交换。
            psv.SetPlotRotation(settings, template.PlotRotation);

            // E. 窗口
            psv.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
            psv.SetPlotWindowArea(settings, new Extents2d(tb.Extents.MinPoint.X, tb.Extents.MinPoint.Y, tb.Extents.MaxPoint.X, tb.Extents.MaxPoint.Y));

            // F. 比例
            psv.SetStdScaleType(settings, StdScaleType.ScaleToFit);
            psv.SetPlotCentered(settings, true);

            // G. 其他
            settings.PrintLineweights = template.PrintLineweights;
            settings.PlotTransparency = template.PlotTransparency;
            settings.ScaleLineweights = template.ScaleLineweights;

            try
            {
                if (template.ShadePlot == PlotSettingsShadePlotType.AsDisplayed || template.ShadePlot == PlotSettingsShadePlotType.Wireframe)
                    settings.ShadePlot = template.ShadePlot;
            }
            catch { }

            psv.RefreshLists(settings);
            info.OverrideSettings = settings;

            return info;
        }

        /// <summary>
        /// 遍历当前打印机支持的所有纸张，寻找尺寸最接近的
        /// </summary>
        private static string FindMatchingMedia(PlotSettings settings, PlotSettingsValidator psv, double width, double height)
        {
            try
            {
                // 获取打印机支持的所有纸张名称 (Canonical Names)
                var mediaList = psv.GetCanonicalMediaNameList(settings);
                if (mediaList == null || mediaList.Count == 0) return null;

                // 定义标准尺寸 (单位: mm)
                // A4=297x210, A3=420x297, A2=594x420, A1=841x594, A0=1189x841
                // 允许误差 5mm
                double tolerance = 5.0;

                // 归一化输入尺寸 (Long x Short)
                double longSide = Math.Max(width, height);
                double shortSide = Math.Min(width, height);

                // 简单的ISO标准判断
                string targetIso = "";
                if (IsSize(longSide, shortSide, 297, 210, tolerance)) targetIso = "A4";
                else if (IsSize(longSide, shortSide, 420, 297, tolerance)) targetIso = "A3";
                else if (IsSize(longSide, shortSide, 594, 420, tolerance)) targetIso = "A2";
                else if (IsSize(longSide, shortSide, 841, 594, tolerance)) targetIso = "A1";
                else if (IsSize(longSide, shortSide, 1189, 841, tolerance)) targetIso = "A0";

                if (string.IsNullOrEmpty(targetIso)) return null; // 非标尺寸，不自动匹配

                // 遍历纸张列表，寻找包含目标ISO名称的纸张 (例如 "ISO_A3_...")
                foreach (string mediaName in mediaList)
                {
                    // 忽略大小写检查
                    if (mediaName.ToUpper().Contains(targetIso))
                    {
                        // 这是一个非常简单的启发式匹配。
                        // 如果要更精确，需要 psv.GetLocaleMediaName(settings, i) 或者解析尺寸
                        // 但通常 "ISO_A3" 或者 "A3" 就够了
                        return mediaName;
                    }
                }
            }
            catch { }
            return null;
        }

        private static bool IsSize(double l1, double s1, double l2, double s2, double tol)
        {
            return Math.Abs(l1 - l2) < tol && Math.Abs(s1 - s2) < tol;
        }

        private static string GetEffectiveName(BlockReference br, Transaction tr)
        {
            if (br == null) return null;
            try
            {
                ObjectId id = br.IsDynamicBlock ? br.DynamicBlockTableRecord : br.BlockTableRecord;
                var btr = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                return btr.Name;
            }
            catch { return null; }
        }

        private static Layout GetLayoutFromBtrId(Transaction tr, Database db, ObjectId btrId)
        {
            var dict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
            foreach (var entry in dict)
            {
                var layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                if (layout.BlockTableRecordId == btrId) return layout;
            }
            return null;
        }
        #endregion
    }
}