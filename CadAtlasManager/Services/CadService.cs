using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using CadAtlasManager.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
// 给 AutoCAD 的 Application 起个别名，防止和 WPF 冲突
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;

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
        // [CadService.cs] 修改后的 BatchPlotByTitleBlocks
        public static List<string> BatchPlotByTitleBlocks(string dwgPath, string outputDir, BatchPlotConfig config)
        {
            List<string> generatedFiles = new List<string>();
            Document doc = null;
            bool isOpenedByUs = false;

            try
            {
                // 1. 获取文档逻辑保持不变
                foreach (Document d in Application.DocumentManager)
                {
                    if (d.Name.Equals(dwgPath, StringComparison.OrdinalIgnoreCase)) { doc = d; break; }
                }
                if (doc == null)
                {
                    if (File.Exists(dwgPath))
                    {
                        doc = Application.DocumentManager.Open(dwgPath, false);
                        isOpenedByUs = true;
                    }
                    else return generatedFiles;
                }

                // 2. 锁定文档并在正确的上下文中执行
                using (doc.LockDocument())
                {
                    Database db = doc.Database;
                    List<string> blockNames = config.TitleBlockNames.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        ObjectId targetSpaceId = db.CurrentSpaceId;
                        var targetLayout = GetLayoutFromBtrId(tr, db, targetSpaceId);

                        var titleBlocks = ScanTitleBlocks(tr, targetSpaceId, blockNames);
                        if (titleBlocks.Count == 0) return generatedFiles;

                        titleBlocks = SortTitleBlocks(titleBlocks, config.OrderType);

                        // 核心改动：将 PlotEngine 移出循环或确保其生命周期极其严格
                        foreach (var tb in titleBlocks)
                        {
                            int index = titleBlocks.IndexOf(tb) + 1;
                            string fileName = Path.GetFileNameWithoutExtension(dwgPath) + (titleBlocks.Count > 1 ? $"_{index:D2}" : "");
                            string fullPdfPath = Path.Combine(outputDir, fileName + ".pdf");

                            // 检查打印机状态
                            if (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting) continue;

                            // 使用 IDisposable 包装所有打印对象
                            using (PlotInfo plotInfo = BuildPlotInfo(tr, targetLayout, tb, config, db))
                            {
                                using (PlotInfoValidator validator = new PlotInfoValidator())
                                {
                                    validator.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                                    validator.Validate(plotInfo);

                                    using (PlotEngine engine = PlotFactory.CreatePublishEngine())
                                    {
                                        engine.BeginPlot(null, null);
                                        engine.BeginDocument(plotInfo, doc.Name, null, 1, true, fullPdfPath);
                                        engine.BeginPage(new PlotPageInfo(), plotInfo, true, null);
                                        engine.BeginGenerateGraphics(null);
                                        engine.EndGenerateGraphics(null);
                                        engine.EndPage(null);
                                        engine.EndDocument(null);
                                        engine.EndPlot(null); // 必须在 engine 销毁前调用
                                    }
                                }
                                generatedFiles.Add(fileName + ".pdf");
                            }
                        }
                        tr.Commit(); // 必须提交事务
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show($"打印核心崩溃: {ex.Message}");
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
        // [CadService.cs] 修改后的 BuildPlotInfo
        private static PlotInfo BuildPlotInfo(Transaction tr, Layout layout, TitleBlockInfo tb, BatchPlotConfig config, Database targetDb)
        {
            PlotInfo info = new PlotInfo();
            info.Layout = layout.ObjectId;

            // 必须使用 using 确保 settings 被释放
            using (PlotSettings settings = new PlotSettings(layout.ModelType))
            {
                settings.CopyFrom(layout);
                var psv = PlotSettingsValidator.Current;

                psv.SetPlotConfigurationName(settings, config.PrinterName, null);

                double blockW = Math.Abs(tb.Extents.MaxPoint.X - tb.Extents.MinPoint.X);
                double blockH = Math.Abs(tb.Extents.MaxPoint.Y - tb.Extents.MinPoint.Y);

                string matchedMedia = FindMatchingMedia(settings, psv, blockW, blockH);
                if (!string.IsNullOrEmpty(matchedMedia))
                    psv.SetCanonicalMediaName(settings, matchedMedia);
                else if (!string.IsNullOrEmpty(config.MediaName))
                    psv.SetCanonicalMediaName(settings, config.MediaName);

                if (!string.IsNullOrEmpty(config.StyleSheet))
                    try { psv.SetCurrentStyleSheet(settings, config.StyleSheet); } catch { }

                psv.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);
                psv.SetPlotWindowArea(settings, new Extents2d(tb.Extents.MinPoint.X, tb.Extents.MinPoint.Y, tb.Extents.MaxPoint.X, tb.Extents.MaxPoint.Y));

                psv.SetPlotRotation(settings, PlotRotation.Degrees000);
                if (config.AutoRotate)
                {
                    double paperW = settings.PlotPaperSize.X;
                    double paperH = settings.PlotPaperSize.Y;
                    if ((blockH > blockW) != (paperH > paperW))
                        psv.SetPlotRotation(settings, PlotRotation.Degrees090);
                }

                if (config.ScaleType == "Fit")
                {
                    psv.SetStdScaleType(settings, StdScaleType.ScaleToFit);
                    psv.SetPlotCentered(settings, true);
                }
                else
                {
                    CustomScale scale = ParseCustomScale(config.ScaleType);
                    psv.SetCustomPrintScale(settings, scale);
                    if (config.PlotCentered) psv.SetPlotCentered(settings, true);
                    else
                    {
                        psv.SetPlotCentered(settings, false);
                        psv.SetPlotOrigin(settings, new Point2d(config.OffsetX, config.OffsetY));
                    }
                }
                psv.RefreshLists(settings);

                // 这一步会将 settings 的内容克隆到 info 中
                info.OverrideSettings = settings;
            } // settings 在这里被 Dispose
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
            var medias = psv.GetCanonicalMediaNameList(s);
            foreach (string mediaName in medias)
            {
                psv.SetCanonicalMediaName(s, mediaName);
                Point2d paperSize = s.PlotPaperSize;
                double pw = paperSize.X; double ph = paperSize.Y;
                if ((Math.Abs(pw - w) < tol && Math.Abs(ph - h) < tol) || (Math.Abs(pw - h) < tol && Math.Abs(ph - w) < tol)) return mediaName;
            }
            return null;
        }
        // [CadService.cs] 提取出来的单页打印核心逻辑
        private static bool PrintExtent(Document doc, Layout layout, TitleBlockInfo tb, BatchPlotConfig config, string fullPdfPath)
        {
            try
            {
                using (PlotEngine engine = PlotFactory.CreatePublishEngine())
                {
                    PlotInfo plotInfo = BuildPlotInfo(null, layout, tb, config, doc.Database); // 注意：BuildPlotInfo 第一个参数传 null 即可
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
                    return true;
                }
            }
            catch { return false; }
        }
        // [CadService.cs]
        public static PlotFileResult EnhancedBatchPlot(string dwgPath, string outputDir, BatchPlotConfig config)
        {
            var result = new PlotFileResult { FilePath = dwgPath, FileName = Path.GetFileName(dwgPath) };
            Document doc = null;
            bool isOpenedByUs = false;

            if (config == null)
            {
                result.IsSuccess = false;
                result.ErrorMessage = "打印配置为空";
                return result;
            }

            try
            {
                // 核心修复：使用 AcApp 明确指定 AutoCAD 的 DocumentManager
                var docMgr = AcApp.DocumentManager;

                // 1. 检查文档是否已经在 CAD 中打开
                foreach (Document d in docMgr)
                {
                    if (d.Name.Equals(dwgPath, StringComparison.OrdinalIgnoreCase))
                    {
                        doc = d;
                        break;
                    }
                }

                // 2. 如果没打开，尝试静默打开
                if (doc == null)
                {
                    if (File.Exists(dwgPath))
                    {
                        doc = docMgr.Open(dwgPath, false); // false 表示不只读，方便打印
                        isOpenedByUs = true;
                    }
                    else
                    {
                        result.IsSuccess = false;
                        result.ErrorMessage = "DWG文件不存在";
                        return result;
                    }
                }

                // 3. 安全检查：如果还是 null，说明打开失败
                if (doc == null)
                {
                    result.IsSuccess = false;
                    result.ErrorMessage = "无法获取 CAD 文档对象";
                    return result;
                }

                // 执行打印流程
                using (doc.LockDocument())
                {
                    // 确保该文档是当前活动文档
                    if (docMgr.MdiActiveDocument != doc)
                        docMgr.MdiActiveDocument = doc;

                    Database db = doc.Database;
                    var blockNames = config.TitleBlockNames.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                                                           .Select(s => s.Trim()).ToList();

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        ObjectId spaceId = db.CurrentSpaceId;
                        var layout = GetLayoutFromBtrId(tr, db, spaceId);
                        var titleBlocks = ScanTitleBlocks(tr, spaceId, blockNames);

                        if (titleBlocks.Count == 0)
                        {
                            result.IsSuccess = false;
                            result.ErrorMessage = "未识别到图框";
                            return result;
                        }

                        titleBlocks = SortTitleBlocks(titleBlocks, config.OrderType);
                        for (int i = 0; i < titleBlocks.Count; i++)
                        {
                            string fileName = Path.GetFileNameWithoutExtension(dwgPath);
                            if (titleBlocks.Count > 1) fileName += $"_{(i + 1):D2}";
                            string fullPdfPath = Path.Combine(outputDir, fileName + ".pdf");

                            // 调用打印逻辑
                            if (PrintExtent(doc, layout, titleBlocks[i], config, fullPdfPath))
                            {
                                result.GeneratedPdfs.Add(fileName + ".pdf");
                            }
                        }
                        result.IsSuccess = true;
                        result.PageCount = result.GeneratedPdfs.Count;
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex) // 明确使用 System.Exception
            {
                result.IsSuccess = false;
                result.ErrorMessage = "异常: " + ex.Message;
            }
            finally
            {
                // 如果是我们打开的文档，打印完关掉
                if (isOpenedByUs && doc != null)
                {
                    doc.CloseAndDiscard();
                }
            }
            return result;
        }
        // [CadService.cs]
        public static int ManualPickAndPlot(Document doc, string outputDir, BatchPlotConfig config)
        {
            var ed = doc.Editor;
            int successCount = 0;

            // 确保 CAD 窗口获得焦点
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = doc;
            doc.Window.Focus();

            using (doc.LockDocument())
            {
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var layout = GetLayoutFromBtrId(tr, doc.Database, doc.Database.CurrentSpaceId);

                    while (true)
                    {
                        // 1. 拾取第一个点
                        var ppr1 = ed.GetPoint("\n指定打印窗口的第一个角点 [直接按回车退出]: ");
                        if (ppr1.Status != PromptStatus.OK) break;

                        // 2. 拾取第二个角点（带矩形框预览）
                        var ppr2 = ed.GetCorner("\n指定第二个角点: ", ppr1.Value);
                        if (ppr2.Status != PromptStatus.OK) break;

                        // 3. 构造虚拟图框范围
                        var tb = new TitleBlockInfo
                        {
                            Extents = new Extents3d(
                                new Point3d(Math.Min(ppr1.Value.X, ppr2.Value.X), Math.Min(ppr1.Value.Y, ppr2.Value.Y), 0),
                                new Point3d(Math.Max(ppr1.Value.X, ppr2.Value.X), Math.Max(ppr1.Value.Y, ppr2.Value.Y), 0)
                            )
                        };

                        // 4. 执行打印
                        string pdfName = $"{Path.GetFileNameWithoutExtension(doc.Name)}_手动_{(successCount + 1):D2}.pdf";
                        if (PrintExtent(doc, layout, tb, config, Path.Combine(outputDir, pdfName)))
                        {
                            successCount++;
                            ed.WriteMessage($"\n已生成第 {successCount} 页手动打印。");
                        }
                    }
                }
            }
            return successCount;
        }
    }
}