using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
// 给 AutoCAD 的 Application 起个别名，防止和 WPF 冲突
using Autodesk.AutoCAD.EditorInput;
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
        // [CadService.cs] 内部添加此方法
        private static void EnsureMetricUnits(Document doc)
        {
            Database db = doc.Database;
            // Measurement: 0 = English (英制), 1 = Metric (公制)
            if (db.Measurement == MeasurementValue.English)
            {
                try
                {
                    // 1. 修改系统变量
                    db.Measurement = MeasurementValue.Metric;

                    // 2. 修改插入单位（可选，但建议一并修改为毫米）
                    db.Insunits = UnitsValue.Millimeters;

                    // 3. 强制保存文件
                    // 注意：打印引擎通常会读取数据库的物理状态，保存可以确保单位变更在打印任务中生效
                    db.SaveAs(doc.Name, db.OriginalFileVersion);

                    doc.Editor.WriteMessage($"\n[系统] 检测到英制图纸，已自动转换为公制并保存：{Path.GetFileName(doc.Name)}");
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\n[错误] 转换公制单位失败: {ex.Message}");
                }
            }
        }
        // [CadService.cs] 增加此方法
        public static void ForceCleanup()
        {
            try
            {
                // 1. 检查打印引擎状态 (虽然 ProcessPlotState 是只读的，但我们可以记录日志或做预防性判断)
                if (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting)
                {
                    // 可以在此处添加日志：警告引擎未完全停止
                }

                // 2. 强制垃圾回收：
                // AutoCAD .NET 包装器在底层是 COM 对象，必须通过两次 GC 确保释放
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForFullGCComplete();

                // 3. 提示：如果依然闪退，说明有静态变量持有了 PlotSettings 或 PlotInfo，
                // 确保没有全局变量在窗口关闭后还引用这些打印对象。
            }
            catch (System.Exception ex)
            {
                // 静默处理清理过程中的异常
                System.Diagnostics.Debug.WriteLine("清理环境时出错: " + ex.Message);
            }
        }

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
        public static int OpenAndManualPlot(string dwgPath, BatchPlotConfig config)
        {
            // 1. 打开并激活文档
            Document doc = Application.DocumentManager.Open(dwgPath, false);
            Application.DocumentManager.MdiActiveDocument = doc;

            // 2. 调用之前写好的 ManualPickAndPlot
            // 它会循环提示用户框选，并使用传入的 config（包含用户刚才在UI改的参数）
            string targetDir = Path.Combine(Path.GetDirectoryName(dwgPath), "_Plot");
            int count = ManualPickAndPlot(doc, targetDir, config);

            // 3. 完成后关闭
            doc.CloseAndDiscard();
            return count;
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

        /// <summary>
        // [CadService.cs]
        public static List<string> BatchPlotByTitleBlocks(string dwgPath, string outputDir, BatchPlotConfig config)
        {
            List<string> generatedFiles = new List<string>();
            Autodesk.AutoCAD.ApplicationServices.Document doc = null;
            bool isOpenedByUs = false;
            var docMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;

            try
            {
                foreach (Autodesk.AutoCAD.ApplicationServices.Document d in docMgr)
                {
                    if (d.Name.Equals(dwgPath, StringComparison.OrdinalIgnoreCase)) { doc = d; break; }
                }

                if (doc == null)
                {
                    if (File.Exists(dwgPath))
                    {
                        doc = docMgr.Open(dwgPath, false);
                        isOpenedByUs = true;
                    }
                    else return generatedFiles;
                }

                using (doc.LockDocument())
                {
                    // --- 新增：打印前确保单位为公制 ---
                    EnsureMetricUnits(doc);
                    // --------------------------------

                    Database db = doc.Database;
                    var blockNames = config.TitleBlockNames?.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries)
                                                           .Select(s => s.Trim()).ToList() ?? new List<string>();

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        ObjectId targetSpaceId = db.CurrentSpaceId;
                        var targetLayout = GetLayoutFromBtrId(tr, db, targetSpaceId);
                        if (targetLayout == null) return generatedFiles;

                        var rawBlocks = ScanTitleBlocks(tr, targetSpaceId, blockNames);
                        if (rawBlocks.Count == 0) return generatedFiles;

                        var titleBlocks = SortTitleBlocks(rawBlocks, config.OrderType);

                        for (int i = 0; i < titleBlocks.Count; i++)
                        {
                            var tb = titleBlocks[i];
                            string fileName = Path.GetFileNameWithoutExtension(dwgPath);

                            // 命名规则：如果有多张图则添加 _01, _02 这种后缀
                            if (titleBlocks.Count > 1)
                                fileName += $"_{(i + 1):D2}";

                            string fullPdfPath = Path.Combine(outputDir, fileName + ".pdf");

                            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                            {
                                if (PrintExtent(doc, targetLayout, tb, config, fullPdfPath))
                                {
                                    generatedFiles.Add(fileName + ".pdf");
                                    PlotMetaManager.SavePlotRecord(dwgPath, fullPdfPath);
                                }
                            }
                        }
                        // 修正：删除了原代码中错误的 result.IsSuccess = true 等行，直接提交事务
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"文件 {Path.GetFileName(dwgPath)} 批量打印核心崩溃: {ex.Message}", ex);
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
        // --- 优化后的 BuildPlotInfo：修复内存崩溃 ---
        // [CadService.cs]
        private static PlotInfo BuildPlotInfo(Transaction tr, Layout layout, TitleBlockInfo tb, BatchPlotConfig config, Database targetDb)
        {
            PlotInfo info = new PlotInfo();
            info.Layout = layout.ObjectId;

            // 【关键修复】：绝对不能在这里使用 using。
            // PlotInfo.OverrideSettings 只是引用了 settings 对象，必须保证在打印完成前它不被释放。
            PlotSettings settings = new PlotSettings(layout.ModelType);
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

            info.OverrideSettings = settings;
            return info;
        }

        // --- 增强的文件检测逻辑 ---
        private static bool WaitForFileAndVerify(string filePath, int timeoutMs)
        {
            int elapsed = 0;
            while (elapsed < timeoutMs)
            {
                if (File.Exists(filePath))
                {
                    try
                    {
                        // 尝试以独占模式打开文件，如果成功说明打印机已放开文件句柄
                        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                        {
                            if (fs.Length > 0) return true;
                        }
                    }
                    catch { /* 文件仍在写入中 */ }
                }
                System.Threading.Thread.Sleep(300);
                elapsed += 300;
            }
            return false;
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

        // --- 1. 新增：等待物理文件生成的辅助方法 ---
        private static bool WaitForFileCreated(string filePath, int timeoutMs)
        {
            int interval = 250; // 每 250 毫秒检查一次
            int elapsed = 0;
            while (elapsed < timeoutMs)
            {
                if (File.Exists(filePath))
                {
                    try
                    {
                        // 关键点：某些 PDF 驱动会先创建一个 0 字节文件占位。
                        // 我们需要确认文件有内容且不再被独占锁定（写入完成）。
                        var info = new FileInfo(filePath);
                        if (info.Length > 0) return true;
                    }
                    catch { /* 文件可能仍在被驱动程序写入并独占锁定，继续等待 */ }
                }
                System.Threading.Thread.Sleep(interval);
                elapsed += interval;
            }
            return false;
        }
        // --- 2. 修改：PrintExtent 核心打印方法 ---
        // --- 修改后的 PrintExtent ---
        // [CadService.cs]
        private static bool PrintExtent(Document doc, Layout layout, TitleBlockInfo tb, BatchPlotConfig config, string fullPdfPath)
        {
            if (PlotFactory.ProcessPlotState != ProcessPlotState.NotPlotting) return false;

            try
            {
                using (PlotEngine engine = PlotFactory.CreatePublishEngine())
                {
                    PlotInfo plotInfo = BuildPlotInfo(null, layout, tb, config, doc.Database);
                    PlotInfoValidator validator = new PlotInfoValidator { MediaMatchingPolicy = MatchingPolicy.MatchEnabled };
                    validator.Validate(plotInfo);

                    engine.BeginPlot(null, null);
                    engine.BeginDocument(plotInfo, doc.Name, null, 1, true, fullPdfPath);
                    engine.BeginPage(new PlotPageInfo(), plotInfo, true, null);
                    engine.BeginGenerateGraphics(null);
                    engine.EndGenerateGraphics(null);
                    engine.EndPage(null);
                    engine.EndDocument(null);
                    engine.EndPlot(null);

                    // 【物理校验】：循环 5 秒等待文件出现且不再被驱动锁定
                    return WaitForFileFinalized(fullPdfPath, 5000);
                }
            }
            catch (System.Exception ex)
            {
                doc.Editor.WriteMessage($"\n[打印错误] {ex.Message}");
                return false;
            }
        }

        private static bool WaitForFileFinalized(string filePath, int timeoutMs)
        {
            int elapsed = 0;
            while (elapsed < timeoutMs)
            {
                if (File.Exists(filePath))
                {
                    try
                    {
                        // 尝试独占打开，成功则说明驱动已放开文件
                        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                        {
                            if (fs.Length > 0) return true;
                        }
                    }
                    catch { }
                }
                System.Threading.Thread.Sleep(300);
                elapsed += 300;
            }
            return false;
        }
        // [CadService.cs]
        /// <summary>
        /// 增强版批量打印：返回带状态的结果对象 (用于 UI 反馈)
        /// </summary>
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
                var docMgr = Application.DocumentManager;
                foreach (Document d in docMgr)
                {
                    if (d.Name.Equals(dwgPath, StringComparison.OrdinalIgnoreCase)) { doc = d; break; }
                }

                if (doc == null)
                {
                    if (File.Exists(dwgPath))
                    {
                        doc = docMgr.Open(dwgPath, false);
                        isOpenedByUs = true;
                    }
                }

                if (doc == null)
                {
                    result.IsSuccess = false;
                    result.ErrorMessage = "无法获取 CAD 文档";
                    return result;
                }

                using (doc.LockDocument())
                {
                    Database db = doc.Database;
                    var blockNames = config.TitleBlockNames.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToList();

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

                            if (PrintExtent(doc, layout, titleBlocks[i], config, fullPdfPath))
                            {
                                // 修正：此处应添加到 result 对象的列表中
                                result.GeneratedPdfs.Add(fileName + ".pdf");
                                PlotMetaManager.SavePlotRecord(dwgPath, fullPdfPath);
                            }
                        }
                        result.IsSuccess = true;
                        result.PageCount = result.GeneratedPdfs.Count;
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                result.IsSuccess = false;
                result.ErrorMessage = "异常: " + ex.Message;
            }
            finally
            {
                if (isOpenedByUs && doc != null) doc.CloseAndDiscard();
            }
            return result;
        }
        // [CadService.cs]
        /// <summary>
        /// 启动 CAD 的外部参照附着对话框（针对 DWG）
        /// </summary>
        public static void OpenXrefDialog()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            doc.Window.Focus();
            // 使用 XATTACH 命令
            doc.SendStringToExecute("_.XATTACH ", true, false, false);
        }

        /// <summary>
        /// 启动 CAD 的图像附着对话框（针对图片）
        /// </summary>
        public static void OpenImageAttachDialog()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            doc.Window.Focus();
            // 使用 IMAGEATTACH 命令，这是专门用于图片的附着窗口
            doc.SendStringToExecute("_.IMAGEATTACH ", true, false, false);
        }
        // [CadService.cs]
        /// <summary>
        /// 启动 CAD 的 PDF 参考底图附着对话框
        /// </summary>
        public static void OpenPdfAttachDialog()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            doc.Window.Focus();
            // 使用 PDFATTACH 命令，专门用于附着 PDF 参考底图
            doc.SendStringToExecute("_.PDFATTACH ", true, false, false);
        }

        // [CadService.cs]

        // [CadService.cs]
        public static int ManualPickAndPlot(Document doc, string outputDir, BatchPlotConfig config)
        {
            var ed = doc.Editor;
            int successCount = 0;
            string dwgPath = doc.Name;

            // 1. 确保 CAD 窗口获得焦点，防止 GetSelection 立即返回
            Autodesk.AutoCAD.ApplicationServices.Application.MainWindow.Focus();
            doc.Window.Focus();

            using (doc.LockDocument())
            {
                // 2. 选择模式
                PromptKeywordOptions pko = new PromptKeywordOptions("\n请选择拾取模式 [手动框选矩形(W) / 框选块参照图框(B)] <W>: ");
                pko.Keywords.Add("W", "W", "手动框选矩形(W)");
                pko.Keywords.Add("B", "B", "框选块参照图框(B)");
                pko.AllowNone = true;

                var pkr = ed.GetKeywords(pko);
                if (pkr.Status != PromptStatus.OK && pkr.Status != PromptStatus.None) return 0;
                string mode = (pkr.Status == PromptStatus.None || pkr.StringResult == "W") ? "W" : "B";

                // 获取当前布局信息（短事务，查完即关）
                ObjectId layoutId;
                using (var tr = doc.Database.TransactionManager.StartTransaction())
                {
                    var btr = tr.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForRead) as BlockTableRecord;
                    layoutId = btr.LayoutId;
                    tr.Commit();
                }

                if (mode == "W")
                {
                    // --- 模式 A：循环手动框选 ---
                    while (true)
                    {
                        var ppr1 = ed.GetPoint("\n指定打印窗口的第一个角点 [Esc结束]: ");
                        if (ppr1.Status != PromptStatus.OK) break;
                        var ppr2 = ed.GetCorner("\n指定第二个角点: ", ppr1.Value);
                        if (ppr2.Status != PromptStatus.OK) break;

                        var tb = new TitleBlockInfo
                        {
                            Extents = new Extents3d(
                                new Point3d(Math.Min(ppr1.Value.X, ppr2.Value.X), Math.Min(ppr1.Value.Y, ppr2.Value.Y), 0),
                                new Point3d(Math.Max(ppr1.Value.X, ppr2.Value.X), Math.Max(ppr1.Value.Y, ppr2.Value.Y), 0)
                            )
                        };

                        using (var tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            var layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                            if (ExecuteSinglePrint(doc, layout, tb, config, outputDir, dwgPath, successCount + 1))
                            {
                                successCount++;
                                PlotMetaManager.SavePlotRecord(dwgPath, Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(dwgPath)}_{successCount:D2}.pdf"));
                            }
                            tr.Commit();
                        }
                    }
                }
                else
                {
                    // --- 模式 B：框选块参照 ---
                    PromptSelectionOptions pso = new PromptSelectionOptions();
                    pso.MessageForAdding = "\n请选择图框对应的块参照或外部参照 (可多选，回车确认选择): ";
                    SelectionFilter filter = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, "INSERT") });

                    var psr = ed.GetSelection(pso, filter);
                    if (psr.Status == PromptStatus.OK)
                    {
                        var pickedBlocks = new List<TitleBlockInfo>();
                        using (var tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            foreach (SelectedObject so in psr.Value)
                            {
                                var br = tr.GetObject(so.ObjectId, OpenMode.ForRead) as BlockReference;
                                if (br != null) try { pickedBlocks.Add(new TitleBlockInfo { BlockName = br.Name, Extents = br.GeometricExtents }); } catch { }
                            }
                            tr.Commit();
                        }

                        var sortedBlocks = SortTitleBlocks(pickedBlocks, config.OrderType);
                        foreach (var tb in sortedBlocks)
                        {
                            using (var tr = doc.Database.TransactionManager.StartTransaction())
                            {
                                var layout = tr.GetObject(layoutId, OpenMode.ForRead) as Layout;
                                if (ExecuteSinglePrint(doc, layout, tb, config, outputDir, dwgPath, successCount + 1))
                                {
                                    successCount++;
                                    PlotMetaManager.SavePlotRecord(dwgPath, Path.Combine(outputDir, $"{Path.GetFileNameWithoutExtension(dwgPath)}_{successCount:D2}.pdf"));
                                }
                                tr.Commit();
                            }
                        }
                    }
                }
            }
            return successCount;
        }
        // 辅助方法：统一执行单页打印操作
        private static bool ExecuteSinglePrint(Document doc, Layout layout, TitleBlockInfo tb, BatchPlotConfig config, string outputDir, string dwgPath, int index)
        {
            string dwgFileName = Path.GetFileNameWithoutExtension(dwgPath);
            string pdfName = $"{dwgFileName}_{index:D2}.pdf";
            string fullPdfPath = Path.Combine(outputDir, pdfName);
            return PrintExtent(doc, layout, tb, config, fullPdfPath);
        }
        // [CadService.cs] 
        // 建议添加在文件末尾或合适的位置
        public static void EnsureHasActiveDocument()
        {
            var docMgr = Application.DocumentManager;
            if (docMgr.Count == 0)
            {
                try
                {
                    // 如果一个文档都没有，则创建一个空白文档（使用默认模板）
                    // 这能防止 PlotSettingsValidator 等 API 因为缺少 Database 而崩溃
                    docMgr.Add("");
                }
                catch
                {
                    // 静默处理
                }
            }
        }
    }
}