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
        /// <summary>
        /// 优化打印环境：关闭后台打印、打印戳记和日志，提升批量打印稳定性
        /// </summary>
        private static void OptimizePlotEnvironment()
        {
            try
            {
                // 1. 关闭后台打印 (0 = 关闭)
                // 这是解决 2021/2024 版本批量打印崩溃、报错“任务已在进行中”的核心设置
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("BACKGROUNDPLOT", 0);

                // 2. 关闭打印戳记 (0 = 关闭)
                // 戳记会触发额外的渲染逻辑，关闭后可提速
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("PLOTSTAMP", 0);

                // 3. 关闭打印日志记录 (0 = 关闭)
                // 防止每打印一张图就写一次 XML/CSV 日志，减少磁盘 I/O 开销
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("PLOTLOG", 0);

                // 4. (可选) 设置 EXPERT 变量，抑制某些打印警告对话框 (1 = 抑制部分提示)
                Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("EXPERT", 1);
            }
            catch (System.Exception ex)
            {
                // 静默处理，防止某些精简版 CAD 不支持某些变量导致崩溃
                System.Diagnostics.Debug.WriteLine("优化环境失败: " + ex.Message);
            }
        }
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

        public static void InsertDwgAsBlock(string filePath)
        {
            if (!File.Exists(filePath)) return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // 聚焦 CAD 窗口
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView();

            try
            {
                using (doc.LockDocument()) // 必须锁定文档
                {
                    string blockName = Path.GetFileNameWithoutExtension(filePath);

                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                        ObjectId btrId = ObjectId.Null;

                        // 检查同名块
                        if (bt.Has(blockName))
                        {
                            // 方案：如果已存在，则直接引用现有块定义，或者你可以选择自动重命名
                            btrId = bt[blockName];
                            ed.WriteMessage($"\n[提示] 块 \"{blockName}\" 已存在，将直接插入现有定义。");
                        }
                        else
                        {
                            // 使用数据库插入 API，这是最稳定的方式
                            using (Database sourceDb = new Database(false, true))
                            {
                                sourceDb.ReadDwgFile(filePath, FileShare.Read, true, "");
                                btrId = db.Insert(blockName, sourceDb, true);
                            }
                        }

                        if (btrId != ObjectId.Null)
                        {
                            // 使用 LISP 的 (command) 函数，这是最稳定的交互式插入方法
                            // _-insert: 强制命令行插入
                            // pause: 等于原来的 \\，让用户点选位置
                            // \"1\" \"1\" \"0\": 分别是 X比例、Y比例、旋转角度
                            string lispCmd = $"(command \"_-insert\" \"{blockName}\" pause \"1\" \"1\" \"0\") ";

                            // 注意：这里第二个参数设为 false，防止自动加空格干扰 LISP 语法
                            doc.SendStringToExecute(lispCmd, false, false, false);
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[错误] 块插入失败: {ex.Message}");
            }
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
            // --- 新增：打印前优化环境 ---
            OptimizePlotEnvironment();
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
                        doc.Editor.WriteMessage($"\n[诊断] 识别到符合条件的图框总数: {titleBlocks.Count}");

                        for (int i = 0; i < titleBlocks.Count; i++)
                        {
                            var tb = titleBlocks[i];
                            string baseName = Path.GetFileNameWithoutExtension(dwgPath);

                            // 【关键修改点】：删除了判断条件，强制为所有导出的 PDF 添加序号后缀
                            // 这样即便只有 1 张图，也会命名为 "文件名_01.pdf"
                            string fileNameWithSuffix = $"{baseName}_{(i + 1):D2}";

                            string fullPdfPath = Path.Combine(outputDir, fileNameWithSuffix + ".pdf");

                            if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                            {
                                // 执行打印
                                if (PrintExtent(doc, targetLayout, tb, config, fullPdfPath, i + 1, titleBlocks.Count))
                                {
                                    // 将带后缀的文件名添加到返回列表
                                    generatedFiles.Add(fileNameWithSuffix + ".pdf");

                                    // 保存打印记录，确保元数据关联的是带后缀的新路径
                                    PlotMetaManager.SavePlotRecord(dwgPath, fullPdfPath);
                                }
                            }
                        }

                        // 提交事务
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
        // [CadService.cs] 修复版：调整赋值顺序，防止 eInvalidInput
        // [CadService.cs] 完整修复版：BuildPlotInfo (需修改调用处传入 Editor)
        private static PlotInfo BuildPlotInfo(Editor ed, Transaction tr, Layout layout, TitleBlockInfo tb, BatchPlotConfig config, Database targetDb)
        {
            PlotInfo info = new PlotInfo();
            info.Layout = layout.ObjectId;

            // 1. 白板模式：创建全新 Settings，不读取原布局
            PlotSettings settings = new PlotSettings(layout.ModelType);
            var psv = PlotSettingsValidator.Current;

            // 2. 建立清洁环境
            try { psv.SetPlotConfigurationName(settings, "None", null); } catch { }

            // 3. 应用打印机
            try
            {
                psv.SetPlotConfigurationName(settings, config.PrinterName, null);
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"无法应用打印机 [{config.PrinterName}]: {ex.Message}");
            }

            // 3.1 强制设置单位为毫米
            try { psv.SetPlotPaperUnits(settings, PlotPaperUnit.Millimeters); } catch { }

            // 4. 应用样式表
            if (!string.IsNullOrEmpty(config.StyleSheet) && config.PlotWithPlotStyles)
            {
                try { psv.SetCurrentStyleSheet(settings, config.StyleSheet); } catch { }
            }

            // 5. 纸张匹配逻辑
            double blockW = Math.Abs(tb.Extents.MaxPoint.X - tb.Extents.MinPoint.X);
            double blockH = Math.Abs(tb.Extents.MaxPoint.Y - tb.Extents.MinPoint.Y);

            if (config.ForceUseSelectedMedia && !string.IsNullOrEmpty(config.MediaName))
            {
                try { psv.SetCanonicalMediaName(settings, config.MediaName); } catch { }
            }
            else
            {
                string matchedMedia = FindMatchingMedia(settings, psv, blockW, blockH);
                if (!string.IsNullOrEmpty(matchedMedia))
                    psv.SetCanonicalMediaName(settings, matchedMedia);
                else if (!string.IsNullOrEmpty(config.MediaName))
                    try { psv.SetCanonicalMediaName(settings, config.MediaName); } catch { }
            }

            // 6. 设置打印区域 (核心修复：坐标转换 + 顺序调整)
            Extents2d window;
            if (layout.ModelType) // 如果是模型空间，必须进行 WCS -> DCS 转换
            {
                try
                {
                    Matrix3d matWcsToDcs = GetWcsToDcsMatrix(ed);
                    Point3d minDcs = tb.Extents.MinPoint.TransformBy(matWcsToDcs);
                    Point3d maxDcs = tb.Extents.MaxPoint.TransformBy(matWcsToDcs);
                    window = new Extents2d(
                        Math.Min(minDcs.X, maxDcs.X),
                        Math.Min(minDcs.Y, maxDcs.Y),
                        Math.Max(minDcs.X, maxDcs.X),
                        Math.Max(minDcs.Y, maxDcs.Y)
                    );
                }
                catch
                {
                    // 转换失败兜底：使用原始坐标
                    window = new Extents2d(
                        new Point2d(tb.Extents.MinPoint.X, tb.Extents.MinPoint.Y),
                        new Point2d(tb.Extents.MaxPoint.X, tb.Extents.MaxPoint.Y)
                    );
                }
            }
            else // 布局空间直接使用图纸坐标
            {
                window = new Extents2d(
                    new Point2d(tb.Extents.MinPoint.X, tb.Extents.MinPoint.Y),
                    new Point2d(tb.Extents.MaxPoint.X, tb.Extents.MaxPoint.Y)
                );
            }

            // 先设置 WindowArea，再设置 PlotType，防止 eInvalidInput
            psv.SetPlotWindowArea(settings, window);
            psv.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);

            // 7. 自动旋转
            psv.SetPlotRotation(settings, Autodesk.AutoCAD.DatabaseServices.PlotRotation.Degrees000);
            if (config.AutoRotate)
            {
                double paperW = settings.PlotPaperSize.X;
                double paperH = settings.PlotPaperSize.Y;
                if ((blockH > blockW) != (paperH > paperW))
                    psv.SetPlotRotation(settings, Autodesk.AutoCAD.DatabaseServices.PlotRotation.Degrees090);
            }

            // 8. 比例与偏移
            psv.SetUseStandardScale(settings, false);
            if (config.ScaleType == "Fit")
            {
                psv.SetStdScaleType(settings, StdScaleType.ScaleToFit);
                psv.SetPlotCentered(settings, true);
            }
            else
            {
                psv.SetCustomPrintScale(settings, ParseCustomScale(config.ScaleType));
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

            // 9. 应用高级选项
            settings.PrintLineweights = config.PlotWithLineweights;
            settings.PlotTransparency = config.PlotTransparency;
            settings.ShowPlotStyles = config.PlotWithPlotStyles;
            settings.ShadePlot = PlotSettingsShadePlotType.AsDisplayed;
            settings.ShadePlotResLevel = ShadePlotResLevel.Normal;

            psv.RefreshLists(settings);
            info.OverrideSettings = settings;
            return info;
        }
        // [CadService.cs] 修正版：获取 WCS 到 DCS 的转换矩阵
        private static Matrix3d GetWcsToDcsMatrix(Editor ed)
        {
            // 【修正点】Editor 没有 ActiveViewport 属性，必须使用 GetCurrentView()
            using (ViewTableRecord vtr = ed.GetCurrentView())
            {
                // 1. 平移：将目标点 (Target) 移到原点
                Matrix3d matDisplace = Matrix3d.Displacement(vtr.Target.GetAsVector().Negate());

                // 2. 投影：将世界坐标系投影到视平面 (World -> Plane)
                // Plane 的法向量是 ViewDirection，原点设为 (0,0,0) 因为我们已经做过平移了
                Matrix3d matPlane = Matrix3d.WorldToPlane(new Plane(Point3d.Origin, vtr.ViewDirection));

                // 3. 扭曲：处理视图扭曲角度 (ViewTwist)
                // 在 DCS 中，Twist 是绕 Z 轴旋转
                Matrix3d matTwist = Matrix3d.Rotation(-vtr.ViewTwist, Vector3d.ZAxis, Point3d.Origin);

                // 组合矩阵：先平移，再投影，最后扭曲 (注意乘法顺序是从右向左应用)
                return matTwist * matPlane * matDisplace;
            }
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

        // [CadService.cs] 用来检测打印过程，看哪一步出错
        private static bool PrintExtent(Document doc, Layout layout, TitleBlockInfo tb, BatchPlotConfig config, string fullPdfPath, int currentIndex, int totalCount)
        {
            var ed = doc.Editor;
            // --- 新增：显示当前进度和图框名称 ---
            ed.WriteMessage($"\n[诊断] === 正在处理图框 ({currentIndex}/{totalCount}): {tb.BlockName} ===");

            // --- 新增：输出详细坐标信息 ---
            ed.WriteMessage($"\n[诊断] 图框物理坐标 (Layout): ");
            ed.WriteMessage($"\n       左下角 (Min): X={tb.Extents.MinPoint.X:F2}, Y={tb.Extents.MinPoint.Y:F2}");
            ed.WriteMessage($"\n       右上角 (Max): X={tb.Extents.MaxPoint.X:F2}, Y={tb.Extents.MaxPoint.Y:F2}");
            ed.WriteMessage($"\n       计算尺寸: 宽={Math.Abs(tb.Extents.MaxPoint.X - tb.Extents.MinPoint.X):F2}, 高={Math.Abs(tb.Extents.MaxPoint.Y - tb.Extents.MinPoint.Y):F2}");

            try
            {
                // 步骤 1: 构建 PlotInfo
                ed.WriteMessage("\n[诊断] 步骤 1: 正在构建 PlotInfo...");

                // 【修改点】：使用 BuildPlotInfoWithLog，只需要传入 ed, layout, tb, config
                PlotInfo plotInfo = BuildPlotInfoWithLog(ed, layout, tb, config);

                ed.WriteMessage("\n[诊断] 步骤 1: PlotInfo 构建成功");

                // 步骤 2: 验证 PlotInfo
                ed.WriteMessage("\n[诊断] 步骤 2: 正在执行 PlotInfoValidator.Validate...");
                PlotInfoValidator validator = new PlotInfoValidator { MediaMatchingPolicy = MatchingPolicy.MatchEnabled };
                validator.Validate(plotInfo);
                ed.WriteMessage("\n[诊断] 步骤 2: 验证通过");

                // 步骤 3: 启动引擎
                ed.WriteMessage("\n[诊断] 步骤 3: 正在创建 PlotEngine...");
                using (PlotEngine engine = PlotFactory.CreatePublishEngine())
                {
                    ed.WriteMessage("\n[诊断] 步骤 4: 正在执行 BeginPlot...");
                    engine.BeginPlot(null, null);

                    ed.WriteMessage($"\n[诊断] 步骤 5: 正在执行 BeginDocument -> {Path.GetFileName(fullPdfPath)}");
                    engine.BeginDocument(plotInfo, doc.Name, null, 1, true, fullPdfPath);

                    ed.WriteMessage("\n[诊断] 步骤 6: 正在执行 BeginPage...");
                    engine.BeginPage(new PlotPageInfo(), plotInfo, true, null);

                    ed.WriteMessage("\n[诊断] 步骤 7: 正在生成图形...");
                    engine.BeginGenerateGraphics(null);
                    engine.EndGenerateGraphics(null);

                    ed.WriteMessage("\n[诊断] 步骤 8: 正在结束页面与文档...");
                    engine.EndPage(null);
                    engine.EndDocument(null);
                    engine.EndPlot(null);
                }

                ed.WriteMessage("\n[诊断] 步骤 9: 物理文件校验中...");
                return WaitForFileFinalized(fullPdfPath, 5000);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[！！！崩溃捕捉] 报错类型: {ex.GetType().Name}");
                ed.WriteMessage($"\n[！！！崩溃捕捉] 错误详情: {ex.Message}");
                ed.WriteMessage($"\n[！！！崩溃捕捉] 堆栈轨迹: \n{ex.StackTrace}");
                return false;
            }
        }

        // [CadService.cs] 修复版 (带日志)
        // [CadService.cs] 完整修复版：BuildPlotInfoWithLog
        private static PlotInfo BuildPlotInfoWithLog(Editor ed, Layout layout, TitleBlockInfo tb, BatchPlotConfig config)
        {
            PlotInfo info = new PlotInfo();
            info.Layout = layout.ObjectId;

            PlotSettings settings = new PlotSettings(layout.ModelType);
            var psv = PlotSettingsValidator.Current;

            ed.WriteMessage($"\n  - [Sub] 重构模式: 正在创建全新打印配置...");

            // 2. 清洁环境
            try { psv.SetPlotConfigurationName(settings, "None", null); } catch { }

            // 3. 应用打印机
            ed.WriteMessage($"\n  - [Sub] 应用打印机: {config.PrinterName}");
            try
            {
                psv.SetPlotConfigurationName(settings, config.PrinterName, null);
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"无法应用打印机 [{config.PrinterName}]: {ex.Message}");
            }

            // 3.1 设置单位
            try { psv.SetPlotPaperUnits(settings, PlotPaperUnit.Millimeters); } catch { }

            // 4. 应用样式表
            if (!string.IsNullOrEmpty(config.StyleSheet) && config.PlotWithPlotStyles)
            {
                try { psv.SetCurrentStyleSheet(settings, config.StyleSheet); } catch { }
            }

            // 5. 纸张匹配
            double blockW = Math.Abs(tb.Extents.MaxPoint.X - tb.Extents.MinPoint.X);
            double blockH = Math.Abs(tb.Extents.MaxPoint.Y - tb.Extents.MinPoint.Y);

            if (config.ForceUseSelectedMedia && !string.IsNullOrEmpty(config.MediaName))
            {
                ed.WriteMessage($"\n  - [Sub] 锁定纸张模式，强制应用: {config.MediaName}");
                try { psv.SetCanonicalMediaName(settings, config.MediaName); }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n  - [警告] 强制应用纸张失败: {ex.Message}，尝试自动匹配...");
                    ApplyAutoMatchMedia(ed, settings, psv, blockW, blockH);
                }
            }
            else
            {
                ApplyAutoMatchMedia(ed, settings, psv, blockW, blockH, config.MediaName);
            }

            // 6. 设置打印区域 (核心修复：坐标转换 + 顺序调整)
            Extents2d window;
            if (layout.ModelType) // 模型空间需转换
            {
                ed.WriteMessage($"\n  - [Sub] 检测到模型空间，正在执行 WCS->DCS 坐标转换...");
                try
                {
                    Matrix3d matWcsToDcs = GetWcsToDcsMatrix(ed);
                    Point3d minDcs = tb.Extents.MinPoint.TransformBy(matWcsToDcs);
                    Point3d maxDcs = tb.Extents.MaxPoint.TransformBy(matWcsToDcs);
                    window = new Extents2d(
                        Math.Min(minDcs.X, maxDcs.X),
                        Math.Min(minDcs.Y, maxDcs.Y),
                        Math.Max(minDcs.X, maxDcs.X),
                        Math.Max(minDcs.Y, maxDcs.Y)
                    );
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n  - [警告] 坐标转换失败: {ex.Message}，将使用原始坐标。");
                    window = new Extents2d(
                        new Point2d(tb.Extents.MinPoint.X, tb.Extents.MinPoint.Y),
                        new Point2d(tb.Extents.MaxPoint.X, tb.Extents.MaxPoint.Y)
                    );
                }
            }
            else // 布局空间
            {
                window = new Extents2d(
                    new Point2d(tb.Extents.MinPoint.X, tb.Extents.MinPoint.Y),
                    new Point2d(tb.Extents.MaxPoint.X, tb.Extents.MaxPoint.Y)
                );
            }

            // 先设置 WindowArea，再设置 PlotType
            psv.SetPlotWindowArea(settings, window);
            psv.SetPlotType(settings, Autodesk.AutoCAD.DatabaseServices.PlotType.Window);

            // 7. 自动旋转
            psv.SetPlotRotation(settings, Autodesk.AutoCAD.DatabaseServices.PlotRotation.Degrees000);
            if (config.AutoRotate)
            {
                double paperW = settings.PlotPaperSize.X;
                double paperH = settings.PlotPaperSize.Y;
                if ((blockH > blockW) != (paperH > paperW))
                {
                    ed.WriteMessage("\n  - [Sub] 智能旋转: 旋转 90 度以适应纸张。");
                    psv.SetPlotRotation(settings, Autodesk.AutoCAD.DatabaseServices.PlotRotation.Degrees090);
                }
            }

            // 8. 比例与偏移
            if (config.ScaleType == "Fit")
            {
                psv.SetStdScaleType(settings, StdScaleType.ScaleToFit);
                psv.SetPlotCentered(settings, true);
            }
            else
            {
                psv.SetCustomPrintScale(settings, ParseCustomScale(config.ScaleType));
                psv.SetPlotCentered(settings, config.PlotCentered);
                if (!config.PlotCentered)
                {
                    psv.SetPlotOrigin(settings, new Point2d(config.OffsetX, config.OffsetY));
                }
            }

            // 9. 高级选项
            settings.PrintLineweights = config.PlotWithLineweights;
            settings.PlotTransparency = config.PlotTransparency;
            settings.ShowPlotStyles = config.PlotWithPlotStyles;
            settings.ShadePlot = PlotSettingsShadePlotType.AsDisplayed;
            settings.ShadePlotResLevel = ShadePlotResLevel.Normal;

            ed.WriteMessage($"\n  - [Sub] 高级选项已应用: 线宽={config.PlotWithLineweights}, 透明度={config.PlotTransparency}");

            psv.RefreshLists(settings);
            info.OverrideSettings = settings;
            return info;
        }

        // 为了代码整洁，提取出来的自动匹配辅助方法 (请将此方法也添加到 CadService 类中)
        private static void ApplyAutoMatchMedia(Editor ed, PlotSettings settings, PlotSettingsValidator psv, double w, double h, string fallbackMedia = null)
        {
            string matchedMedia = FindMatchingMedia(settings, psv, w, h);

            if (!string.IsNullOrEmpty(matchedMedia))
            {
                ed.WriteMessage($"\n  - [Sub] 自动匹配到标准纸张: {matchedMedia}");
                psv.SetCanonicalMediaName(settings, matchedMedia);
            }
            else if (!string.IsNullOrEmpty(fallbackMedia))
            {
                try
                {
                    ed.WriteMessage($"\n  - [Sub] 未匹配到标准纸张，正在尝试应用预设纸张: {fallbackMedia}");
                    psv.SetCanonicalMediaName(settings, fallbackMedia);
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\n  - [警告] 预设纸张无效: {ex.Message}");
                }
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

        /// <summary>
        /// 增强版批量打印：返回带状态的结果对象 (用于 UI 反馈)
        /// </summary>
        public static PlotFileResult EnhancedBatchPlot(string dwgPath, string outputDir, BatchPlotConfig config)
        {
            // --- 新增：打印前优化环境 ---
            OptimizePlotEnvironment();
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
                            // 1. 获取不带后缀的基础文件名
                            string baseName = Path.GetFileNameWithoutExtension(dwgPath);

                            // 2. 强制生成带序号的文件名 (例如: 文件名_01.pdf)
                            // 注意：这里将变量名统一，防止出现 fileNameWithoutExt 未定义的错误
                            string finalPdfName = $"{baseName}_{(i + 1):D2}.pdf";

                            // 3. 使用【带序号】的文件名合成完整路径
                            string fullPdfPath = Path.Combine(outputDir, finalPdfName);

                            // 4. 执行打印，传入正确的 fullPdfPath
                            if (PrintExtent(doc, layout, titleBlocks[i], config, fullPdfPath, i + 1, titleBlocks.Count))
                            {
                                // 5. 修正：将【带序号】的文件名添加到结果列表
                                result.GeneratedPdfs.Add(finalPdfName);

                                // 6. 保存记录，确保指纹关联的是带序号的 PDF
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
            return PrintExtent(doc, layout, tb, config, fullPdfPath, index, index);
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