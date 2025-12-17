using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections; // 必须引用
using System.IO;

namespace CadAtlasManager
{
    public class CadService : IExtensionApplication
    {
        private const string FINGERPRINT_KEY = "CadAtlas_SmartFingerprint";

        // =================================================================
        // 【第一部分】插件初始化与保存监听
        // =================================================================

        public void Initialize()
        {
            Application.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;
            foreach (Document doc in Application.DocumentManager)
            {
                SubscribeToSaveEvent(doc);
            }
        }

        public void Terminate() { }

        private void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            SubscribeToSaveEvent(e.Document);
        }

        private void SubscribeToSaveEvent(Document doc)
        {
            if (doc == null) return;
            doc.Database.BeginSave += Database_BeginSave;
        }

        private void Database_BeginSave(object sender, DatabaseIOEventArgs e)
        {
            Database db = sender as Database;
            if (db == null) return;

            try
            {
                // 1. 判断内容修改 (DBMOD 第0位)
                object dbmodObj = Application.GetSystemVariable("DBMOD");
                int dbmod = Convert.ToInt32(dbmodObj);
                bool hasContentChanged = (dbmod & 1) == 1;

                // 2. 读取当前指纹
                string currentFingerprint = GetSmartFingerprintFromDb(db);

                // 3. 更新指纹
                if (hasContentChanged || string.IsNullOrEmpty(currentFingerprint))
                {
                    UpdateSmartFingerprint(db);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("写入指纹失败: " + ex.Message);
            }
        }

        // 【修复 1】写入指纹：像操作普通字典一样操作 CustomPropertyTable
        private void UpdateSmartFingerprint(Database db)
        {
            try
            {
                DatabaseSummaryInfoBuilder infoBuilder = new DatabaseSummaryInfoBuilder(db.SummaryInfo);
                IDictionary props = infoBuilder.CustomPropertyTable;

                string newFingerprint = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                if (props.Contains(FINGERPRINT_KEY))
                {
                    props[FINGERPRINT_KEY] = newFingerprint; // 更新
                }
                else
                {
                    props.Add(FINGERPRINT_KEY, newFingerprint); // 新增
                }

                db.SummaryInfo = infoBuilder.ToDatabaseSummaryInfo();
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("UpdateSmartFingerprint Error: " + ex.Message);
            }
        }

        // 【修复 2】读取指纹：使用 while 循环遍历迭代器，解决 foreach 报错
        private string GetSmartFingerprintFromDb(Database db)
        {
            try
            {
                DatabaseSummaryInfo info = db.SummaryInfo;
                IDictionaryEnumerator iter = info.CustomProperties; // 获取迭代器

                // IDictionaryEnumerator 必须用 while 遍历
                while (iter.MoveNext())
                {
                    if (iter.Key.ToString() == FINGERPRINT_KEY)
                    {
                        return iter.Value.ToString();
                    }
                }
                return null;
            }
            catch { return null; }
        }

        // =================================================================
        // 【第二部分】对外静态方法 (供 AtlasView 调用)
        // =================================================================

        public static string GetSmartFingerprint(string dwgPath)
        {
            if (!File.Exists(dwgPath)) return "0";
            try
            {
                using (Database db = new Database(false, true))
                {
                    db.ReadDwgFile(dwgPath, FileShare.Read, false, null);

                    // 1. 尝试读取自定义属性 (使用 while 循环)
                    DatabaseSummaryInfo info = db.SummaryInfo;
                    IDictionaryEnumerator iter = info.CustomProperties;

                    while (iter.MoveNext())
                    {
                        if (iter.Key.ToString() == FINGERPRINT_KEY)
                        {
                            return iter.Value.ToString();
                        }
                    }

                    // 2. 降级方案
                    return db.Tdupdate.ToString("o");
                }
            }
            catch
            {
                return "0";
            }
        }

        public static PlotSettings GetTemplatePlotSettings()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return null;

            PlotSettings template = null;
            try
            {
                using (DocumentLock docLock = doc.LockDocument())
                {
                    using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        LayoutManager layoutMgr = LayoutManager.Current;
                        ObjectId layoutId = layoutMgr.GetLayoutId(layoutMgr.CurrentLayout);
                        Layout layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);

                        template = new PlotSettings(layout.ModelType);
                        template.CopyFrom(layout);
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex) { System.Windows.MessageBox.Show(ex.Message); }
            return template;
        }

        public static bool PlotToPdfWithSettings(string dwgPath, string outputPdfPath, PlotSettings templateSettings)
        {
            bool success = false;
            Document doc = null;
            bool needsClose = false;

            try
            {
                foreach (Document d in Application.DocumentManager)
                {
                    if (d.Name.Equals(dwgPath, StringComparison.OrdinalIgnoreCase)) { doc = d; break; }
                }
                if (doc == null)
                {
                    doc = Application.DocumentManager.Open(dwgPath, false);
                    needsClose = true;
                }

                using (DocumentLock docLock = doc.LockDocument())
                {
                    Database db = doc.Database;
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        LayoutManager layoutMgr = LayoutManager.Current;
                        string layoutName = "Model";
                        ObjectId layoutId = layoutMgr.GetLayoutId(layoutName);
                        if (layoutId.IsNull) layoutId = layoutMgr.GetLayoutId(layoutMgr.CurrentLayout);

                        Layout layout = (Layout)tr.GetObject(layoutId, OpenMode.ForRead);
                        PlotInfo plotInfo = new PlotInfo();
                        plotInfo.Layout = layout.ObjectId;

                        PlotSettings newSettings = new PlotSettings(layout.ModelType);
                        newSettings.CopyFrom(templateSettings);

                        PlotSettingsValidator psv = PlotSettingsValidator.Current;

                        // 明确使用 DatabaseServices 下的 PlotType
                        if (newSettings.PlotType == Autodesk.AutoCAD.DatabaseServices.PlotType.Extents)
                        {
                            psv.RefreshLists(newSettings);
                        }

                        plotInfo.OverrideSettings = newSettings;
                        PlotInfoValidator validator = new PlotInfoValidator();
                        validator.MediaMatchingPolicy = MatchingPolicy.MatchEnabled;
                        validator.Validate(plotInfo);

                        if (PlotFactory.ProcessPlotState == ProcessPlotState.NotPlotting)
                        {
                            using (PlotEngine engine = PlotFactory.CreatePublishEngine())
                            {
                                engine.BeginPlot(null, null);
                                engine.BeginDocument(plotInfo, doc.Name, null, 1, true, outputPdfPath);
                                PlotPageInfo pageInfo = new PlotPageInfo();
                                engine.BeginPage(pageInfo, plotInfo, true, null);
                                engine.BeginGenerateGraphics(null);
                                engine.EndGenerateGraphics(null);
                                engine.EndPage(null);
                                engine.EndDocument(null);
                                engine.EndPlot(null);
                            }
                            success = true;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
                success = false;
            }
            finally
            {
                if (needsClose && doc != null) doc.CloseAndDiscard();
            }
            return success;
        }

        // =================================================================
        // 【第三部分】基础操作
        // =================================================================

        public static void OpenDwg(string sourcePath, string mode, string targetCopyPath = null)
        {
            if (!File.Exists(sourcePath)) return;
            DocumentCollection acDocMgr = Application.DocumentManager;
            string finalOpenPath = sourcePath;
            bool isReadOnly = mode == "Read";
            if (mode == "Copy" && !string.IsNullOrEmpty(targetCopyPath)) { File.Copy(sourcePath, targetCopyPath, true); finalOpenPath = targetCopyPath; isReadOnly = false; }
            foreach (Document doc in acDocMgr) { if (doc.Name.Equals(finalOpenPath, StringComparison.OrdinalIgnoreCase)) { acDocMgr.MdiActiveDocument = doc; return; } }
            acDocMgr.Open(finalOpenPath, isReadOnly);
        }

        public static void InsertDwgAsBlock(string filePath)
        {
            if (!File.Exists(filePath)) return;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            try
            {
                using (DocumentLock l = doc.LockDocument())
                {
                    using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        // 简单插入逻辑
                        doc.Editor.Command("_.INSERT", filePath, "0,0", "1", "1", "0");
                        tr.Commit();
                    }
                }
            }
            catch { }
        }
    }
}