// 【关键修改：引用新拆分的命名空间】
using Autodesk.AutoCAD.ApplicationServices;
using CadAtlasManager.Models;
using CadAtlasManager.UI;
// 添加这两个引用
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression; // 引用 System.IO.Compression.FileSystem
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using UserControl = System.Windows.Controls.UserControl;
using WinForms = System.Windows.Forms;

namespace CadAtlasManager
{
    public partial class AtlasView : UserControl
    {
        public ObservableCollection<FileSystemItem> Items { get; set; }
        public ObservableCollection<FileSystemItem> ProjectTreeItems { get; set; }
        public ObservableCollection<FileSystemItem> PlotFolderItems { get; set; }
        public ObservableCollection<ProjectItem> ProjectList { get; set; }
        // 新增：PlotFileListItems 用于右侧列表
        public ObservableCollection<FileSystemItem> PlotFileListItems { get; set; }
        private List<string> _loadedAtlasFolders = new List<string>();
        private ProjectItem _activeProject = null;
        private readonly string _versionInfo = "CAD图集管理器 v3.6 (Refactored)\n\n更新：\n1. 代码结构重构优化\n2. 准备接入新功能";

        private Point _dragStartPoint;
        private FileSystemItem _draggedItem;

        // 多选与备注相关
        private FileSystemItem _lastSelectedItem = null;
        private FileSystemItem _currentRemarkItem = null; // 当前正在写备注的文件

        private readonly List<string> _allowedExtensions = new List<string>
        {
            ".dwg", ".dxf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".wps", ".pdf", ".txt",
            ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".mp4", ".avi", ".mov",
            ".zip", ".rar", ".7z"
        };

        public AtlasView()
        {
            Items = new ObservableCollection<FileSystemItem>();
            ProjectTreeItems = new ObservableCollection<FileSystemItem>();
            PlotFolderItems = new ObservableCollection<FileSystemItem>();
            PlotFileListItems = new ObservableCollection<FileSystemItem>();
            ProjectList = new ObservableCollection<ProjectItem>();

            InitializeComponent();

            FileTree.ItemsSource = Items;
            ProjectTree.ItemsSource = ProjectTreeItems;

            // 修改绑定
            PlotFolderTree.ItemsSource = PlotFolderItems;
            PlotFileList.ItemsSource = PlotFileListItems;

            CbProjects.ItemsSource = ProjectList;
            LoadConfig();
        }
        private void RefreshPlotTree()
        {
            PlotFolderItems.Clear();
            PlotFileListItems.Clear();

            if (_activeProject == null) return;

            // 1. 确保基础路径存在
            if (string.IsNullOrEmpty(_activeProject.OutputPath))
                _activeProject.OutputPath = Path.Combine(_activeProject.Path, "_Plot");

            if (!Directory.Exists(_activeProject.OutputPath))
            {
                try { Directory.CreateDirectory(_activeProject.OutputPath); } catch { }
            }
            if (!Directory.Exists(_activeProject.OutputPath)) return;

            // 2. 创建“成果 PDF”目录 (物理隔离，管理更清晰)
            string combinedPath = Path.Combine(_activeProject.OutputPath, "Combined");
            if (!Directory.Exists(combinedPath)) Directory.CreateDirectory(combinedPath);

            // =========================================================
            // 构建左侧树：两个固定节点
            // =========================================================

            // A. 节点：分项 PDF (对应 _Plot 根目录)
            var itemSplit = CreateItem(_activeProject.OutputPath, ExplorerItemType.Folder, true);
            itemSplit.Name = "📄 分项 PDF"; // 改名
            itemSplit.TypeIcon = "📂";

            // 加载子文件夹 (但要排除 Combined 文件夹，防止重复显示)
            LoadPlotFoldersOnly(itemSplit, "Combined");
            PlotFolderItems.Add(itemSplit);

            // B. 节点：成果 PDF (对应 _Plot\Combined 子目录)
            var itemCombined = CreateItem(combinedPath, ExplorerItemType.Folder, true);
            itemCombined.Name = "📑 成果 PDF"; // 新增项
            itemCombined.TypeIcon = "📚";

            // 成果目录下通常不需要再分级，当然也可以加载
            LoadPlotFoldersOnly(itemCombined);
            PlotFolderItems.Add(itemCombined);

            // =========================================================
            // ✅ 自动加载逻辑
            // =========================================================
            // 默认展开“分项 PDF”并选中
            itemSplit.IsExpanded = true;
            itemSplit.IsItemSelected = true; // 设定 UI 选中状态

            // 强制加载“分项 PDF”下的文件到右侧列表
            LoadPlotFilesList(itemSplit);
        }

        // 修改方法签名，增加 excludeName 参数，默认由 null
        private void LoadPlotFoldersOnly(FileSystemItem parent, string excludeName = null)
        {
            try
            {
                foreach (var dir in Directory.GetDirectories(parent.FullPath))
                {
                    var dirInfo = new DirectoryInfo(dir);

                    // 跳过隐藏文件夹
                    if (dirInfo.Attributes.HasFlag(FileAttributes.Hidden)) continue;

                    // ✅ 新增过滤：如果文件夹名等于我们要排除的名字（比如 Combined），则跳过
                    if (!string.IsNullOrEmpty(excludeName) &&
                        dirInfo.Name.Equals(excludeName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var sub = CreateItem(dir, ExplorerItemType.Folder);
                    LoadPlotFoldersOnly(sub); // 递归加载
                    parent.Children.Add(sub);
                }
            }
            catch { }
        }

        // 当左侧树选择变化时，加载右侧文件列表
        private void PlotFolderTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var folder = e.NewValue as FileSystemItem;
            LoadPlotFilesList(folder);
        }

        private void LoadPlotFilesList(FileSystemItem folder)
        {
            PlotFileListItems.Clear();
            if (folder == null || !Directory.Exists(folder.FullPath)) return;

            try
            {
                // 加载该文件夹下的所有文件 (不递归)
                foreach (var file in Directory.GetFiles(folder.FullPath))
                {
                    string ext = Path.GetExtension(file).ToLower();
                    // 仅显示 PDF, 图片等
                    if (".pdf.jpg.jpeg.png.plt".Contains(ext))
                    {
                        var item = CreateItem(file, ExplorerItemType.File);

                        // 填充日期
                        item.CreationDate = File.GetCreationTime(file).ToString("yyyy-MM-dd HH:mm");

                        // 预先检查一下版本状态（如果是PDF）
                        if (ext == ".pdf") ValidatePdfVersion(item);

                        PlotFileListItems.Add(item);
                    }
                }
            }
            catch { }
        }

        // 全选/反选 按钮点击
        private void OnSelectAllPlotFiles_Click(object sender, RoutedEventArgs e)
        {
            var cb = sender as CheckBox;
            bool check = cb.IsChecked == true;
            foreach (var item in PlotFileListItems)
            {
                item.IsChecked = check;
            }
        }

        // 双击列表文件打开
        private void PlotFileList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (PlotFileList.SelectedItem is FileSystemItem item)
            {
                Process.Start(new ProcessStartInfo(item.FullPath) { UseShellExecute = true });
            }
        }

        // 占位方法：删除
        // =================================================================
        // 【Phase 2 新增】删除与合并具体实现
        // =================================================================

        // 1. 删除功能
        private void BtnDeletePlotFiles_Click(object sender, RoutedEventArgs e)
        {
            // 获取所有勾选的项目
            var targets = PlotFileListItems.Where(i => i.IsChecked).ToList();

            if (targets.Count == 0)
            {
                MessageBox.Show("请先在列表中勾选要删除的文件。", "提示");
                return;
            }

            // 弹出确认框
            var result = MessageBox.Show($"确定要永久删除这 {targets.Count} 个文件吗？\n此操作无法撤销。",
                                         "删除确认", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes) return;

            int successCount = 0;
            try
            {
                foreach (var item in targets)
                {
                    if (File.Exists(item.FullPath))
                    {
                        File.Delete(item.FullPath);

                        // 同步删除备注
                        RemarkManager.HandleDelete(item.FullPath);
                        successCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除过程中发生错误：\n{ex.Message}", "错误");
            }

            // 刷新当前列表
            // 获取当前选中的文件夹节点，重新加载它的内容
            var currentFolder = GetSelectedItem();
            if (currentFolder != null && currentFolder.Type == ExplorerItemType.Folder)
            {
                LoadPlotFilesList(currentFolder);
            }
            else
            {
                // 如果找不到当前节点，就刷新整个树
                RefreshPlotTree();
            }

            if (successCount > 0)
            {
                // 可以在这里加个简单的提示，或者直接静默
            }
        }

        // 2. 合并 PDF 功能
        private void BtnMergePdf_Click(object sender, RoutedEventArgs e)
        {
            // 1. 获取勾选的 PDF 文件
            // 注意：我们按照文件名排序，确保合并顺序符合直觉
            var targets = PlotFileListItems
                .Where(i => i.IsChecked && i.FullPath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                .OrderBy(i => i.Name)
                .ToList();

            if (targets.Count < 2)
            {
                MessageBox.Show("请至少勾选 2 个 PDF 文件进行合并。", "提示");
                return;
            }

            // 2. 弹出重命名框 (复用已有的 RenameDialog)
            // 默认文件名：取第一个文件的名字 + "_合并"
            string defaultName = Path.GetFileNameWithoutExtension(targets[0].Name) + "_合并";
            var dlg = new RenameDialog(defaultName);
            dlg.Title = "合并 PDF - 输入新文件名";

            if (dlg.ShowDialog() != true) return;

            // 3. 确定保存路径
            // 为了管理规范，我们将所有合并成果统一存放到 "Combined" 文件夹
            string saveDir = Path.Combine(_activeProject.OutputPath, "Combined");
            if (!Directory.Exists(saveDir)) Directory.CreateDirectory(saveDir);

            string savePath = Path.Combine(saveDir, dlg.NewName + ".pdf");

            // 检查重名
            if (File.Exists(savePath))
            {
                if (MessageBox.Show($"文件 {dlg.NewName}.pdf 已存在，是否覆盖？", "覆盖确认",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                    return;
            }

            // 4. 执行合并
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                // 收集源文件路径
                List<string> sourceFiles = targets.Select(t => t.FullPath).ToList();

                // 调用合并核心方法
                MergePdfFiles(sourceFiles, savePath);

                MessageBox.Show($"成功合并 {targets.Count} 个文件！\n已保存至：成果 PDF 目录", "成功");

                // 5. 刷新界面
                // 既然文件保存到了 "Combined" (成果 PDF)，我们应该刷新整个树，让用户能看到新文件
                RefreshPlotTree();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"合并失败：\n{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        // 3. PdfSharp 合并核心逻辑
        private void MergePdfFiles(List<string> sourceFiles, string destFile)
        {
            // 创建一个新的 PDF 文档
            using (PdfDocument outputDocument = new PdfDocument())
            {
                foreach (string file in sourceFiles)
                {
                    // 以导入模式打开文档
                    using (PdfDocument inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import))
                    {
                        // 遍历每一页并添加到输出文档
                        int count = inputDocument.PageCount;
                        for (int idx = 0; idx < count; idx++)
                        {
                            // 获取页
                            PdfPage page = inputDocument.Pages[idx];
                            // 添加到新文档
                            outputDocument.AddPage(page);
                        }
                    }
                }
                // 保存
                outputDocument.Save(destFile);
            }

            // TODO: (Phase 3) 这里未来将添加“版本校验”的元数据写入逻辑
            // PlotMetaManager.SaveCombinedRecord(destFile, sourceFiles);
        }
        // ... [保留原有的 CreateItem, ValidatePdfVersion 等方法，
        // 注意 ValidatePdfVersion 不需要改动，因为它操作的是 FileSystemItem 对象，
        // 我们的 DataGrid 绑定的也是 FileSystemItem，所以状态更新会自动反映在表格中] ...


        // =================================================================
        // 【核心：备注面板动画与逻辑】
        // =================================================================

        private void MenuItem_Remark_Click(object sender, RoutedEventArgs e)
        {
            var item = GetSelectedItem();
            if (item != null)
            {
                if (!item.IsItemSelected) item.IsItemSelected = true;
                ShowRemarkPanel(item);
                TbRemark.Focus(); // 聚焦输入
            }
        }

        private void ShowRemarkPanel(FileSystemItem item)
        {
            if (item == null) return;
            _currentRemarkItem = item;

            // 加载备注内容
            string remark = RemarkManager.GetRemark(item.FullPath);

            // 填充并展开
            TbRemark.TextChanged -= TbRemark_TextChanged;
            TbRemark.Text = remark ?? "";
            TbRemark.TextChanged += TbRemark_TextChanged;

            // 动画
            DoubleAnimation ani = new DoubleAnimation(150, TimeSpan.FromMilliseconds(250));
            ani.EasingFunction = new CubicEase { EasingMode = EasingMode.EaseOut };
            RemarkPanel.BeginAnimation(HeightProperty, ani);
        }

        private void HideRemarkPanel()
        {
            if (RemarkPanel.Height == 0) return;

            DoubleAnimation ani = new DoubleAnimation(0, TimeSpan.FromMilliseconds(200));
            ani.EasingFunction = new CubicEase { EasingMode = EasingMode.EaseIn };
            RemarkPanel.BeginAnimation(HeightProperty, ani);
            _currentRemarkItem = null;
        }

        // 文本变化：实时保存
        private void TbRemark_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_currentRemarkItem != null)
            {
                string text = TbRemark.Text.Trim();
                RemarkManager.SaveRemark(_currentRemarkItem.FullPath, text);
                _currentRemarkItem.HasRemark = !string.IsNullOrEmpty(text);
            }
        }

        // 选中变化：如果点的不是正在编辑的文件，收起面板
        private void OnTreeSelectionChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var newItem = e.NewValue as FileSystemItem;
            if (RemarkPanel.Height > 0)
            {
                if (newItem == null || newItem != _currentRemarkItem)
                {
                    HideRemarkPanel();
                }
            }
        }

        private void BtnCloseRemark_Click(object sender, RoutedEventArgs e) => HideRemarkPanel();

        // =================================================================
        // 【文件加载逻辑】(已集成备注读取)
        // =================================================================

        private void ReloadAtlasTree()
        {
            Items.Clear();
            foreach (var path in _loadedAtlasFolders) if (Directory.Exists(path)) AddFolderToTree(path);
        }

        private void AddFolderToTree(string path)
        {
            // 加载该文件夹的备注数据到内存
            RemarkManager.LoadRemarks(path);

            var rootItem = CreateItem(path, ExplorerItemType.Folder, true);
            if (LoadSubItems(rootItem) || string.IsNullOrEmpty(TbSearch.Text)) Items.Add(rootItem);
        }

        private bool LoadSubItems(FileSystemItem parent)
        {
            bool match = false;
            string search = TbSearch.Text.Trim().ToLower();
            string filter = (CbFileType.SelectedItem as ComboBoxItem)?.Content.ToString();
            bool isSearching = !string.IsNullOrEmpty(search);

            // 加载子文件夹时，也要预加载备注
            RemarkManager.LoadRemarks(parent.FullPath);

            try
            {
                foreach (var dir in Directory.GetDirectories(parent.FullPath))
                {
                    // 跳过隐藏文件夹 (比如 .cadatlas)
                    if (new DirectoryInfo(dir).Attributes.HasFlag(FileAttributes.Hidden)) continue;

                    var sub = CreateItem(dir, ExplorerItemType.Folder);
                    bool childMatch = LoadSubItems(sub);
                    if (childMatch || (isSearching && sub.Name.ToLower().Contains(search)) || !isSearching)
                    {
                        parent.Children.Add(sub);
                        match = true;
                        if (isSearching) { parent.IsExpanded = true; sub.IsExpanded = childMatch; }
                    }
                }
                foreach (var file in Directory.GetFiles(parent.FullPath))
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (_allowedExtensions.Contains(ext) && CheckFileFilter(ext, filter))
                    {
                        if (!isSearching || Path.GetFileName(file).ToLower().Contains(search))
                        {
                            parent.Children.Add(CreateItem(file, ExplorerItemType.File));
                            match = true;
                            if (isSearching) parent.IsExpanded = true;
                        }
                    }
                }
            }
            catch { }
            return match;
        }

        private void RefreshProjectTree()
        {
            ProjectTreeItems.Clear();
            if (_activeProject == null || !Directory.Exists(_activeProject.Path)) return;

            RemarkManager.LoadRemarks(_activeProject.Path);

            var root = CreateItem(_activeProject.Path, ExplorerItemType.Folder, true);
            root.Name = _activeProject.Name; // 显示项目别名
            root.TypeIcon = "🏗️";

            LoadProjectSubItems(root);
            ProjectTreeItems.Add(root);
        }

        private void LoadProjectSubItems(FileSystemItem parent)
        {
            try
            {
                RemarkManager.LoadRemarks(parent.FullPath); // 预加载
                string filter = (CbProjectFileType.SelectedItem as ComboBoxItem)?.Content.ToString();

                foreach (var dir in Directory.GetDirectories(parent.FullPath))
                {
                    if (new DirectoryInfo(dir).Attributes.HasFlag(FileAttributes.Hidden)) continue;

                    var sub = CreateItem(dir, ExplorerItemType.Folder);
                    LoadProjectSubItems(sub);
                    if (filter == "所有格式" || sub.Children.Count > 0) parent.Children.Add(sub);
                }
                foreach (var file in Directory.GetFiles(parent.FullPath))
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (_allowedExtensions.Contains(ext) && CheckFileFilter(ext, filter))
                    {
                        parent.Children.Add(CreateItem(file, ExplorerItemType.File));
                    }
                }
            }
            catch { }
        }


        // =================================================================
        // 【重构】批量打印流程 (智能筛选 + 统一参数 + 防崩溃修复)
        // =================================================================

        private void MenuItem_BatchPlot_Click(object sender, RoutedEventArgs e)
        {
            // 1. 基础检查
            if (_activeProject == null)
            {
                MessageBox.Show("请先在项目工作台中选择一个项目。");
                return;
            }

            // 2. 获取需要打印的 DWG 文件
            var selectedItems = GetAllSelectedItems();
            // 智能容错：如果没有多选，但当前选中了一个文件，则当作单选处理
            if (selectedItems.Count == 0 && GetSelectedItem() != null)
            {
                selectedItems.Add(GetSelectedItem());
            }

            var dwgFiles = selectedItems
                .Where(i => i.Type == ExplorerItemType.File && i.FullPath.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (dwgFiles.Count == 0)
            {
                MessageBox.Show("请选择至少一个 DWG 图纸文件。");
                return;
            }

            // =================================================================
            // 【核心修复】防止“零文档状态”导致崩溃
            // 如果当前 CAD 没有打开任何图纸，打印引擎会初始化失败导致闪退。
            // 这里自动新建一个空白文档作为“垫底”，确保环境安全。
            // =================================================================
            if (Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Count == 0)
            {
                try
                {
                    // Add(null) 会使用默认模板新建一个 Drawing1.dwg
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.Add(null);
                }
                catch
                {
                    // 忽略新建失败的情况，继续尝试打印，避免阻断流程
                }
            }
            // =================================================================

            // 3. 准备输出目录 (_Plot)
            string outputDir = Path.Combine(_activeProject.Path, "_Plot");
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            // 加载历史记录 (用于判断版本状态)
            PlotMetaManager.LoadHistory(outputDir);

            // 4. 构建候选列表 (数据准备)
            var candidates = new List<PlotCandidate>();
            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                foreach (var item in dwgFiles)
                {
                    string dwgName = item.Name;
                    string pdfName = Path.GetFileNameWithoutExtension(dwgName) + ".pdf";

                    // 1. 获取当前时间戳 (快速)
                    string currentTimestamp = CadService.GetFileTimestamp(item.FullPath);

                    // 2. 智能校验 (传入委托，按需读取 Tduupdate)
                    bool isLatest = PlotMetaManager.CheckStatus(outputDir, pdfName, dwgName, currentTimestamp,
                        () => CadService.GetContentFingerprint(item.FullPath)); // 委托

                    candidates.Add(new PlotCandidate
                    {
                        FileName = dwgName,
                        FilePath = item.FullPath,
                        IsOutdated = !isLatest,       // 注意取反
                        IsSelected = !isLatest,       // 默认勾选过期的
                        NewTdUpdate = currentTimestamp, // 这里暂存时间戳
                        VersionStatus = !isLatest ? "⚠️ 需更新" : "✅ 最新"
                    });
                }
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }

            // 5. 【关键】弹出我们新建的 BatchPlotDialog
            // 此时 CadAtlasManager.UI 命名空间下已经是新的 XAML 窗口了
            var dialog = new BatchPlotDialog(candidates);

            // 如果用户点了取消，或者关闭了窗口，直接返回
            if (dialog.ShowDialog() != true) return;

            // 6. 获取用户设置的参数
            var config = dialog.FinalConfig;        // 打印机、纸张、样式等配置
            var filesToPrint = dialog.ConfirmedFiles; // 用户最终勾选的文件

            if (filesToPrint.Count == 0) return;

            // 7. 开始批量打印循环
            int totalSuccess = 0;
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                foreach (var filePath in filesToPrint)
                {
                    // 【修改】获取生成的 PDF 列表 (List<string>)
                    List<string> generatedPdfs = CadService.BatchPlotByTitleBlocks(filePath, outputDir, config);

                    if (generatedPdfs.Count > 0)
                    {
                        // 打印成功，立即获取该 DWG 的最新指纹 (Tduupdate) 和 时间戳
                        string freshFingerprint = CadService.GetContentFingerprint(filePath);
                        string freshTime = CadService.GetFileTimestamp(filePath);
                        string dwgName = Path.GetFileName(filePath); // 记录文件名即可，因为校验时有递归搜索

                        // 【关键】为每一个生成的 PDF 都保存一条记录
                        // 这样 Drawing1_1.pdf 也能关联到 Drawing1.dwg
                        foreach (var pdfName in generatedPdfs)
                        {
                            PlotMetaManager.SaveRecord(outputDir, pdfName, dwgName, freshFingerprint, freshTime);
                        }

                        totalSuccess += generatedPdfs.Count;
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"批处理过程中发生意外错误:\n{ex.Message}");
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }

            // 8. 刷新界面
            RefreshPlotTree(); // 刷新输出目录树

            string msg = $"批量打印完成！\n共处理文件: {filesToPrint.Count} 个\n生成 PDF 页数: {totalSuccess} 页";
            MessageBox.Show(msg, "完成");
        }

        private void LoadPlotSubItems(FileSystemItem parent)
        {
            try
            {
                RemarkManager.LoadRemarks(parent.FullPath);
                // 仅加载 PDF, JPG, PNG, PLT
                string plotExts = ".pdf.jpg.jpeg.png.plt";

                foreach (var dir in Directory.GetDirectories(parent.FullPath))
                {
                    if (new DirectoryInfo(dir).Attributes.HasFlag(FileAttributes.Hidden)) continue;
                    var sub = CreateItem(dir, ExplorerItemType.Folder);
                    LoadPlotSubItems(sub);
                    parent.Children.Add(sub);
                }
                foreach (var file in Directory.GetFiles(parent.FullPath))
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (plotExts.Contains(ext))
                    {
                        parent.Children.Add(CreateItem(file, ExplorerItemType.File));
                    }
                }
            }
            catch { }
        }

        // 统一创建对象的方法，处理备注状态
        private FileSystemItem CreateItem(string path, ExplorerItemType type, bool isRoot = false)
        {
            string ext = Path.GetExtension(path).ToLower();
            string remark = RemarkManager.GetRemark(path);

            // FileSystemItem 现在位于 CadAtlasManager.Models 命名空间
            return new FileSystemItem
            {
                Name = Path.GetFileName(path),
                FullPath = path,
                Type = type,
                TypeIcon = type == ExplorerItemType.Folder ? "📁" : GetIconForExtension(ext),
                IsRoot = isRoot,
                IsExpanded = isRoot,
                FontWeight = isRoot ? FontWeights.Bold : FontWeights.Normal,
                HasRemark = !string.IsNullOrEmpty(remark)
            };
        }

        // =================================================================
        // 【文件操作同步逻辑】
        // =================================================================

        private void MenuItem_Rename_Click(object sender, RoutedEventArgs e)
        {
            var item = GetSelectedItem(); if (item == null) return;
            string dir = Path.GetDirectoryName(item.FullPath);
            string nameNoExt = item.Type == ExplorerItemType.File ? Path.GetFileNameWithoutExtension(item.FullPath) : Path.GetFileName(item.FullPath);

            // 假设 RenameDialog 仍然是一个独立的 XAML 窗口文件
            var dlg = new RenameDialog(nameNoExt);

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    string newName = item.Type == ExplorerItemType.File ? dlg.NewName + Path.GetExtension(item.FullPath) : dlg.NewName;
                    string newPath = Path.Combine(dir, newName);

                    if (item.Type == ExplorerItemType.File) File.Move(item.FullPath, newPath);
                    else Directory.Move(item.FullPath, newPath);

                    // 【同步备注】
                    RemarkManager.HandleRename(item.FullPath, newPath);

                    ReloadAtlasTree(); RefreshProjectTree(); RefreshPlotTree();
                }
                catch (System.Exception ex) { MessageBox.Show("重命名失败: " + ex.Message); }
            }
        }

        private void MenuItem_Delete_Click(object sender, RoutedEventArgs e)
        {
            var items = GetAllSelectedItems();
            if (items.Count == 0 && GetSelectedItem() != null) items.Add(GetSelectedItem());
            if (items.Count == 0) return;

            if (MessageBox.Show($"确定删除选中的 {items.Count} 个项目吗？", "确认", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                try
                {
                    foreach (var item in items)
                    {
                        if (item.Type == ExplorerItemType.File) File.Delete(item.FullPath);
                        else Directory.Delete(item.FullPath, true);

                        // 【同步删除备注】
                        RemarkManager.HandleDelete(item.FullPath);
                    }
                    RefreshProjectTree(); RefreshPlotTree();
                }
                catch (System.Exception ex) { MessageBox.Show("删除失败: " + ex.Message); }
            }
        }

        private void ProjectTree_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var targetTreeViewItem = GetTreeViewItemUnderMouse(e.OriginalSource as DependencyObject);
                FileSystemItem targetItem = targetTreeViewItem?.DataContext as FileSystemItem;

                if (targetItem == null || targetItem.Type != ExplorerItemType.Folder) return;

                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files == null) return;

                foreach (var file in files)
                {
                    if (Path.GetDirectoryName(file) == targetItem.FullPath) continue;
                    try
                    {
                        string destPath = Path.Combine(targetItem.FullPath, Path.GetFileName(file));
                        if (File.Exists(file)) File.Move(file, destPath);
                        else if (Directory.Exists(file)) Directory.Move(file, destPath);

                        // 【同步移动备注】
                        RemarkManager.HandleMove(file, destPath);
                    }
                    catch (Exception ex) { MessageBox.Show($"移动 {Path.GetFileName(file)} 失败: " + ex.Message); }
                }
                RefreshProjectTree();
            }
        }

        // ========================= 其他基础逻辑 =========================

        // [修改文件: UI/AtlasView.xaml.cs]

        private void OnTreeItemClick(object sender, MouseButtonEventArgs e)
        {
            // 1. 【关键修改】必须调用这个方法，才能穿透文字找到背后的数据
            var clickedItem = GetClickedItem(e);

            // 如果没点到任何项（比如点到了空白处），直接返回
            if (clickedItem == null) return;

            var treeView = sender as TreeView;
            var allItems = treeView.ItemsSource as ObservableCollection<FileSystemItem>;

            // 2. 处理多选 (Ctrl)
            if (Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl))
            {
                clickedItem.IsItemSelected = !clickedItem.IsItemSelected; // 反选
                _lastSelectedItem = clickedItem;
            }
            // 3. 处理连选 (Shift)
            else if (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift))
            {
                if (_lastSelectedItem != null && allItems != null)
                    SelectRange(allItems, _lastSelectedItem, clickedItem);
            }
            // 4. 处理单选
            else
            {
                // 排除点击箭头（折叠/展开）的情况
                var originalObj = e.OriginalSource as DependencyObject;
                bool isArrow = false;
                while (originalObj != null && !(originalObj is TreeViewItem))
                {
                    if (originalObj is System.Windows.Controls.Primitives.ToggleButton)
                    {
                        isArrow = true;
                        break;
                    }
                    originalObj = VisualTreeHelper.GetParent(originalObj);
                }

                // 只有点在非箭头区域，才触发单选逻辑
                if (!isArrow)
                {
                    // 如果它还没被选中，或者为了清除其他多选项 -> 执行单选
                    // (注意：这里不要加 !clickedItem.IsItemSelected 判断，否则单机已选项无法清除其他多选项)
                    ClearAllSelection(Items);
                    ClearAllSelection(ProjectTreeItems);
                    ClearAllSelection(PlotFolderItems);
                    foreach (var item in PlotFileListItems) item.IsChecked = false; // 清除列表勾选

                    clickedItem.IsItemSelected = true;
                    _lastSelectedItem = clickedItem;
                }
            }
        }

        // =================================================================
        // 【核心修改】复制移动功能 (单选重命名 / 多选批量备份)
        // =================================================================
        private void MenuItem_CopyInPlace_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = GetAllSelectedItems();

            if (selectedItems.Count == 0)
            {
                var current = GetSelectedItem();
                if (current != null) selectedItems.Add(current);
            }

            var filesToCopy = selectedItems.Where(i => i.Type == ExplorerItemType.File).ToList();

            if (filesToCopy.Count == 0)
            {
                MessageBox.Show("请选择至少一个文件进行复制。");
                return;
            }

            string defaultPath = Path.GetDirectoryName(filesToCopy[0].FullPath);
            string defaultName = filesToCopy.Count == 1 ? filesToCopy[0].Name : "";

            // 使用自定义弹窗 (CopyMoveDialog - 现在位于 CadAtlasManager.UI 命名空间)
            var dlg = new CopyMoveDialog(defaultPath, defaultName, filesToCopy.Count > 1);
            dlg.Title = filesToCopy.Count == 1 ? $"复制移动 - {filesToCopy[0].Name}" : $"批量复制移动 ({filesToCopy.Count}个文件)";

            if (dlg.ShowDialog() == true)
            {
                string targetDir = dlg.SelectedPath;
                int successCount = 0;

                try
                {
                    if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);

                    if (filesToCopy.Count == 1)
                    {
                        // 单选：支持重命名
                        string sourceFile = filesToCopy[0].FullPath;
                        string newFileName = dlg.FileName;
                        string destPath = Path.Combine(targetDir, newFileName);

                        if (File.Exists(destPath))
                        {
                            if (MessageBox.Show($"文件 {newFileName} 已存在，是否覆盖？", "冲突", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                                return;
                        }

                        File.Copy(sourceFile, destPath, true);

                        // 同步备注
                        string remark = RemarkManager.GetRemark(sourceFile);
                        if (!string.IsNullOrEmpty(remark)) RemarkManager.SaveRemark(destPath, remark);

                        successCount++;
                    }
                    else
                    {
                        // 多选：批量复制
                        foreach (var item in filesToCopy)
                        {
                            string destPath = Path.Combine(targetDir, item.Name);
                            if (File.Exists(destPath)) continue;

                            File.Copy(item.FullPath, destPath, true);

                            string remark = RemarkManager.GetRemark(item.FullPath);
                            if (!string.IsNullOrEmpty(remark)) RemarkManager.SaveRemark(destPath, remark);

                            successCount++;
                        }
                    }

                    RefreshProjectTree();
                    if (successCount > 0) MessageBox.Show($"成功复制 {successCount} 个文件到：\n{targetDir}");
                }
                catch (System.Exception ex) { MessageBox.Show("复制失败: " + ex.Message); }
            }
        }

        // =================================================================
        // 【修改】一键打包功能 (使用自定义 ZipSaveDialog)
        // =================================================================
        private void MenuItem_Zip_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = GetAllSelectedItems();
            var contextItem = GetSelectedItem();
            if (contextItem != null && !contextItem.IsItemSelected)
            {
                selectedItems.Clear();
                selectedItems.Add(contextItem);
            }

            if (selectedItems.Count == 0)
            {
                MessageBox.Show("请先选择要打包的文件或文件夹。");
                return;
            }

            string defaultPath = Path.GetDirectoryName(selectedItems[0].FullPath);
            string defaultName = $"打包_{DateTime.Now:yyyyMMdd_HHmm}.zip";

            // 使用自定义弹窗 (ZipSaveDialog - 现在位于 CadAtlasManager.UI 命名空间)
            var dlg = new ZipSaveDialog(defaultPath, defaultName);

            if (dlg.ShowDialog() == true)
            {
                string targetDir = dlg.SelectedPath;
                string fileName = dlg.FileName;

                if (!fileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase)) fileName += ".zip";
                string zipFullPath = Path.Combine(targetDir, fileName);

                try
                {
                    if (File.Exists(zipFullPath))
                    {
                        if (MessageBox.Show($"文件 {fileName} 已存在，是否覆盖？", "确认", MessageBoxButton.YesNo) != MessageBoxResult.Yes)
                            return;
                        File.Delete(zipFullPath);
                    }

                    using (ZipArchive archive = ZipFile.Open(zipFullPath, ZipArchiveMode.Create))
                    {
                        foreach (var item in selectedItems)
                        {
                            if (item.Type == ExplorerItemType.File) archive.CreateEntryFromFile(item.FullPath, item.Name);
                            else AddFolderToZip(archive, item.FullPath, item.Name);
                        }
                    }

                    if (MessageBox.Show("打包成功！是否打开所在文件夹？", "完成", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        Process.Start("explorer.exe", "/select,\"" + zipFullPath + "\"");
                    }
                }
                catch (System.Exception ex) { MessageBox.Show("打包失败: " + ex.Message); }
            }
        }

        private void AddFolderToZip(ZipArchive archive, string sourceDir, string entryPrefix)
        {
            foreach (string file in Directory.GetFiles(sourceDir))
                archive.CreateEntryFromFile(file, Path.Combine(entryPrefix, Path.GetFileName(file)));
            foreach (string dir in Directory.GetDirectories(sourceDir))
                AddFolderToZip(archive, dir, Path.Combine(entryPrefix, Path.GetFileName(dir)));
        }

        private void MenuItem_InsertBlock_Click(object sender, RoutedEventArgs e)
        {
            var items = GetAllSelectedItems();
            if (items.Count == 0 && GetSelectedItem() != null) items.Add(GetSelectedItem());
            foreach (var item in items) if (item.Type == ExplorerItemType.File && (item.FullPath.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))) CadService.InsertDwgAsBlock(item.FullPath);
        }

        private void LoadConfig()
        {
            try
            {
                var config = ConfigManager.Load(); if (config == null) return;
                if (config.AtlasFolders != null) foreach (var f in config.AtlasFolders) if (Directory.Exists(f) && !_loadedAtlasFolders.Contains(f)) { _loadedAtlasFolders.Add(f); AddFolderToTree(f); }
                if (config.Projects != null) foreach (var p in config.Projects) if (Directory.Exists(p.Path)) ProjectList.Add(p);
                if (!string.IsNullOrEmpty(config.LastActiveProjectPath)) { var t = ProjectList.FirstOrDefault(p => p.Path == config.LastActiveProjectPath); if (t != null) CbProjects.SelectedItem = t; }
            }
            catch { }
        }
        private void SaveConfig() { string ap = _activeProject?.Path ?? ""; ConfigManager.Save(_loadedAtlasFolders, ProjectList, ap); }

        private void BtnLoadFolder_Click(object sender, RoutedEventArgs e) { using (var d = new WinForms.FolderBrowserDialog()) { if (d.ShowDialog() == WinForms.DialogResult.OK && !_loadedAtlasFolders.Contains(d.SelectedPath)) { _loadedAtlasFolders.Add(d.SelectedPath); AddFolderToTree(d.SelectedPath); SaveConfig(); } } }
        private void BtnAddProject_Click(object sender, RoutedEventArgs e) { using (var d = new WinForms.FolderBrowserDialog()) { if (d.ShowDialog() == WinForms.DialogResult.OK && !ProjectList.Any(p => p.Path == d.SelectedPath)) { var p = new ProjectItem(Path.GetFileName(d.SelectedPath), d.SelectedPath); ProjectList.Add(p); CbProjects.SelectedItem = p; SaveConfig(); } } }

        private void ClearAllSelection(ObservableCollection<FileSystemItem> items) { if (items == null) return; foreach (var i in items) { i.IsItemSelected = false; ClearAllSelection(i.Children); } }
        private List<FileSystemItem> GetAllSelectedItems()
        {
            var l = new List<FileSystemItem>();
            CollectSelected(Items, l);
            CollectSelected(ProjectTreeItems, l);

            // --- 修改开始 ---
            CollectSelected(PlotFolderItems, l);   // 收集左侧树选中项

            // 收集右侧列表的勾选项 (注意：列表用的是 IsChecked 属性，或者 DataGrid 的 SelectedItems，这里兼容处理)
            // 假设我们主要依赖复选框 IsChecked
            foreach (var item in PlotFileListItems)
            {
                if (item.IsChecked || item.IsItemSelected) l.Add(item);
            }
            // --- 修改结束 ---

            return l;
        }
        private void CollectSelected(ObservableCollection<FileSystemItem> s, List<FileSystemItem> r) { if (s == null) return; foreach (var i in s) { if (i.IsItemSelected) r.Add(i); CollectSelected(i.Children, r); } }
        private void SelectRange(ObservableCollection<FileSystemItem> root, FileSystemItem s, FileSystemItem e) { var l = new List<FileSystemItem>(); FlattenTree(root, l); int i1 = l.IndexOf(s), i2 = l.IndexOf(e); if (i1 != -1 && i2 != -1) for (int i = Math.Min(i1, i2); i <= Math.Max(i1, i2); i++) l[i].IsItemSelected = true; }
        private void FlattenTree(ObservableCollection<FileSystemItem> n, List<FileSystemItem> r) { foreach (var node in n) { r.Add(node); if (node.IsExpanded) FlattenTree(node.Children, r); } }
        private TreeViewItem GetTreeViewItemUnderMouse(DependencyObject e) { while (e != null && !(e is TreeViewItem)) e = VisualTreeHelper.GetParent(e); return e as TreeViewItem; }
        private void MenuItem_CopyInPlace_Click_Legacy(object sender, RoutedEventArgs e) { /* 保留旧逻辑引用防止报错，实际使用新的 CopyMove */ }
        private void MenuItem_Remove_Click(object sender, RoutedEventArgs e) { if (FileTree.SelectedItem is FileSystemItem i && i.IsRoot) { _loadedAtlasFolders.Remove(i.FullPath); Items.Remove(i); SaveConfig(); } }
        private void BtnRemoveProject_Click(object sender, RoutedEventArgs e)
        {
            if (CbProjects.SelectedItem is ProjectItem p)
            {
                ProjectList.Remove(p);
                ProjectTreeItems.Clear();

                // --- 修改开始 ---
                PlotFolderItems.Clear();   // 清空文件夹树
                PlotFileListItems.Clear(); // 清空文件列表
                                           // --- 修改结束 ---

                _activeProject = null;
                SaveConfig();
            }
        }
        private void CbProjects_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CbProjects.SelectedItem is ProjectItem p)
            {
                _activeProject = p;
                RefreshProjectTree();
                RefreshPlotTree(); // 【关键】同时刷新图纸工作台
                SaveConfig();
            }
        }
        private void BtnRefreshProject_Click(object sender, RoutedEventArgs e) => RefreshProjectTree();
        private void BtnRefresh_Click(object sender, RoutedEventArgs e) => ReloadAtlasTree();
        private void BtnRefreshPlot_Click(object sender, RoutedEventArgs e) => RefreshPlotTree();
        private void TbSearch_TextChanged(object sender, TextChangedEventArgs e) => ReloadAtlasTree();
        private void CbFileType_SelectionChanged(object sender, SelectionChangedEventArgs e) => ReloadAtlasTree();
        private void CbProjectFileType_SelectionChanged(object sender, SelectionChangedEventArgs e) => RefreshProjectTree();
        private void BtnOpenProjectFolder_Click(object sender, RoutedEventArgs e) { if (_activeProject != null) Process.Start("explorer.exe", _activeProject.Path); }
        private void BtnOpenPlotFolder_Click(object sender, RoutedEventArgs e) { if (_activeProject != null && Directory.Exists(_activeProject.OutputPath)) Process.Start("explorer.exe", _activeProject.OutputPath); }

        // =================================================================
        // 【新增】图纸版本校验逻辑
        // =================================================================

        // 1. 点击校验按钮
        private void BtnCheckVersion_Click(object sender, RoutedEventArgs e)
        {
            if (_activeProject == null) return;
            string plotDir = _activeProject.OutputPath;
            if (string.IsNullOrEmpty(plotDir) || !Directory.Exists(plotDir)) return;

            PlotMetaManager.LoadHistory(plotDir);

            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                // --- 修改开始 ---
                // 只校验当前右侧列表中的文件
                CheckItemsVersion(PlotFileListItems);
                // --- 修改结束 ---
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        // 递归检查树节点
        private void CheckItemsVersion(ObservableCollection<FileSystemItem> nodes)
        {
            foreach (var item in nodes)
            {
                if (item.Type == ExplorerItemType.File)
                {
                    // 仅检查 PDF 文件
                    if (item.FullPath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                    {
                        ValidatePdfVersion(item);
                    }
                }
                else
                {
                    // 递归检查子文件夹
                    if (item.Children.Count > 0) CheckItemsVersion(item.Children);
                }
            }
        }

        /// =========================================================
        // 【核心修改】文件校验逻辑
        // =========================================================
        private void ValidatePdfVersion(FileSystemItem pdfItem)
        {
            string pdfName = pdfItem.Name;

            // 1. 优先从历史记录反查 DWG 文件名 (包含相对路径信息)
            string realDwgName = PlotMetaManager.GetSourceDwgName(_activeProject.OutputPath, pdfName);

            // 2. 如果历史记录查不到，尝试智能猜测 (去掉 _1, _2 后缀)
            if (string.IsNullOrEmpty(realDwgName))
            {
                string baseName = System.Text.RegularExpressions.Regex.Replace(Path.GetFileNameWithoutExtension(pdfName), @"_\d+$", "");
                realDwgName = baseName + ".dwg";
            }

            // 3. 构建源文件路径 (先尝试直接在项目根目录找)
            string sourceDwgPath = Path.Combine(_activeProject.Path, realDwgName);

            // 【关键修复】如果根目录找不到，进行全目录递归搜索 (解决子文件夹文件丢失问题)
            if (!File.Exists(sourceDwgPath))
            {
                // 仅搜索文件名
                string onlyFileName = Path.GetFileName(realDwgName);
                try
                {
                    // 在项目目录下递归搜索同名文件
                    var foundFiles = Directory.GetFiles(_activeProject.Path, onlyFileName, SearchOption.AllDirectories);
                    if (foundFiles.Length > 0)
                    {
                        // 找到第一个匹配的
                        sourceDwgPath = foundFiles[0];
                        // 更新一下 realDwgName 为文件名，以便后续逻辑统一
                        realDwgName = Path.GetFileName(sourceDwgPath);
                    }
                }
                catch { } // 忽略权限错误等
            }

            // 4. 最终检查文件是否存在
            if (!File.Exists(sourceDwgPath))
            {
                pdfItem.VersionStatus = "❓ 源文件丢失";
                pdfItem.StatusColor = Brushes.Gray;
                return;
            }

            // 5. 校验状态 (传入 Tduupdate 获取委托)
            // 获取当前源文件的时间戳 (文件系统)
            string currentTimestamp = CadService.GetFileTimestamp(sourceDwgPath);

            // 调用 Manager 进行比对
            bool isLatest = PlotMetaManager.CheckStatus(_activeProject.OutputPath, pdfName, realDwgName, currentTimestamp,
                () => CadService.GetContentFingerprint(sourceDwgPath)); // 这里会读取 Tduupdate

            if (!isLatest)
            {
                pdfItem.VersionStatus = "⚠️ 已过期";
                pdfItem.StatusColor = Brushes.Red;
            }
            else
            {
                pdfItem.VersionStatus = "✅ 最新";
                pdfItem.StatusColor = Brushes.Green;
            }
        }

        // =================================================================
        // 【核心修复】右键菜单 - 打开源 DWG (同步最新的查找逻辑)
        // =================================================================
        private void MenuItem_OpenSourceDwg_Click(object sender, RoutedEventArgs e)
        {
            // --- 修改开始 ---
            // 从右侧列表获取选中项
            var item = PlotFileList.SelectedItem as FileSystemItem;
            if (item == null || item.Type != ExplorerItemType.File || !item.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("请先选择一个 PDF 文件。");
                return;
            }

            if (_activeProject == null) return;

            string pdfName = item.Name;
            string outputDir = _activeProject.OutputPath; // PDF 所在的 _Plot 目录

            // ---------------------------------------------------------
            // 查找逻辑 (与校验版本逻辑保持一致)
            // ---------------------------------------------------------

            // A. 优先从历史记录反查真实 DWG 文件名
            string realDwgName = PlotMetaManager.GetSourceDwgName(outputDir, pdfName);

            // B. 如果记录不存在，尝试智能猜测 (去除 _1, _2 等后缀)
            if (string.IsNullOrEmpty(realDwgName))
            {
                // 正则去除结尾的 _数字： Drawing1_1.pdf -> Drawing1.dwg
                string baseName = System.Text.RegularExpressions.Regex.Replace(Path.GetFileNameWithoutExtension(pdfName), @"_\d+$", "");
                realDwgName = baseName + ".dwg";
            }

            // C. 构建路径 & 递归搜索 (解决子文件夹问题)
            string sourceDwgPath = Path.Combine(_activeProject.Path, realDwgName);

            if (!File.Exists(sourceDwgPath))
            {
                try
                {
                    // 在项目根目录下递归搜索同名文件
                    string onlyFileName = Path.GetFileName(realDwgName);
                    var foundFiles = Directory.GetFiles(_activeProject.Path, onlyFileName, SearchOption.AllDirectories);
                    if (foundFiles.Length > 0)
                    {
                        sourceDwgPath = foundFiles[0]; // 找到了！
                    }
                }
                catch { }
            }

            // ---------------------------------------------------------
            // 执行打开
            // ---------------------------------------------------------
            if (File.Exists(sourceDwgPath))
            {
                // "Edit" 表示以可读写方式打开，方便你修改源文件
                CadService.OpenDwg(sourceDwgPath, "Edit");
            }
            else
            {
                MessageBox.Show($"未找到源文件：\n{realDwgName}\n\n可能原因：\n1. 源文件已被移动或重命名\n2. 尚未对此文件进行过批量打印(无关联记录)", "无法打开");
            }
        }
        private void BtnVersion_Click(object sender, RoutedEventArgs e) => MessageBox.Show(_versionInfo);
        private void ProjectTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) { _dragStartPoint = e.GetPosition(null); var t = GetTreeViewItemUnderMouse(e.OriginalSource as DependencyObject); if (t != null) _draggedItem = t.DataContext as FileSystemItem; OnTreeItemClick(sender, e); }
        private void ProjectTree_MouseMove(object sender, MouseEventArgs e) { if (e.LeftButton == MouseButtonState.Pressed && _draggedItem != null) { var diff = _dragStartPoint - e.GetPosition(null); if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance || Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance) { var s = GetAllSelectedItems(); if (!s.Contains(_draggedItem)) { s.Clear(); s.Add(_draggedItem); } DataObject d = new DataObject(DataFormats.FileDrop, s.Select(i => i.FullPath).ToArray()); DragDrop.DoDragDrop(ProjectTree, d, DragDropEffects.Copy | DragDropEffects.Move); _draggedItem = null; } } }
        private void MenuItem_Open_Click(object sender, RoutedEventArgs e) { var i = GetSelectedItem(); if (i != null && i.Type == ExplorerItemType.File) ExecuteFileAction(i); }
        private FileSystemItem GetSelectedItem()
        {
            if (ProjectTree.IsVisible && ProjectTree.SelectedItem != null) return ProjectTree.SelectedItem as FileSystemItem;

            // --- 修改开始 ---
            // 优先检查右侧文件列表的选中项
            if (PlotFileList.SelectedItem != null) return PlotFileList.SelectedItem as FileSystemItem;
            // 其次检查左侧文件夹树的选中项
            if (PlotFolderTree.SelectedItem != null) return PlotFolderTree.SelectedItem as FileSystemItem;
            // --- 修改结束 ---

            if (_lastSelectedItem != null && _lastSelectedItem.IsItemSelected) return _lastSelectedItem;
            return FileTree.SelectedItem as FileSystemItem;
        }
        private void MenuItem_NewFolder_Click(object sender, RoutedEventArgs e) { var i = GetSelectedItem(); if (i != null && i.Type == ExplorerItemType.Folder) { Directory.CreateDirectory(Path.Combine(i.FullPath, "新建文件夹")); RefreshProjectTree(); } }
        private void MenuItem_Expand_Click(object sender, RoutedEventArgs e) { var i = GetSelectedItem(); if (i != null) SetExpansion(i, true); }
        private void MenuItem_Collapse_Click(object sender, RoutedEventArgs e) { var i = GetSelectedItem(); if (i != null) SetExpansion(i, false); }
        private void SetExpansion(FileSystemItem i, bool e) { i.IsExpanded = e; foreach (var c in i.Children) if (c.Type == ExplorerItemType.Folder) SetExpansion(c, e); }
        private void FileTree_MouseDoubleClick(object sender, MouseButtonEventArgs e) { if (GetClickedItem(e) is FileSystemItem i && i.Type == ExplorerItemType.File) ExecuteFileAction(i); }
        private void ProjectTree_MouseDoubleClick(object sender, MouseButtonEventArgs e) { if (GetClickedItem(e) is FileSystemItem i && i.Type == ExplorerItemType.File) OpenFileSmart(i.FullPath, "Edit"); }
        private void PlotTree_MouseDoubleClick(object sender, MouseButtonEventArgs e) { if (GetClickedItem(e) is FileSystemItem i && i.Type == ExplorerItemType.File) Process.Start(new ProcessStartInfo(i.FullPath) { UseShellExecute = true }); }

        // [修改文件: UI/AtlasView.xaml.cs] 请将此方法添加到类中

        // [修改文件: UI/AtlasView.xaml.cs]

        private void OnTreeItemRightClick(object sender, MouseButtonEventArgs e)
        {
            var clickedItem = GetClickedItem(e);
            if (clickedItem != null)
            {
                // 1. 【核心修复】如果点击的项已经是选中状态，则保留多选，什么都不做
                // 这样右键菜单就会基于当前的多选列表弹出来
                if (clickedItem.IsItemSelected)
                {
                    return;
                }

                // 2. 如果点击的项没被选中，则按标准流程：清除旧选择 -> 选中当前项
                ClearAllSelection(Items);
                ClearAllSelection(ProjectTreeItems);
                ClearAllSelection(PlotFolderItems);
                foreach (var item in PlotFileListItems) item.IsChecked = false; // 清除列表勾选

                clickedItem.IsItemSelected = true;
                _lastSelectedItem = clickedItem;
            }
        }
        private FileSystemItem GetClickedItem(MouseButtonEventArgs e)
        {
            // 获取点击位置的原始元素
            DependencyObject obj = e.OriginalSource as DependencyObject;

            // 向上遍历可视化树，直到找到 TreeViewItem
            while (obj != null && !(obj is TreeViewItem))
            {
                // 处理 Run (文字内容) 等非 Visual 元素
                if (obj is System.Windows.FrameworkContentElement fce)
                {
                    obj = fce.Parent;
                }
                else
                {
                    obj = VisualTreeHelper.GetParent(obj);
                }
            }

            // 如果找到了 TreeViewItem，返回其绑定的数据
            return (obj as TreeViewItem)?.DataContext as FileSystemItem;
        }
        // [修改文件: UI/AtlasView.xaml.cs]

        private void ExecuteFileAction(FileSystemItem i)
        {
            // 1. 确定操作模式
            string mode = "Read";
            if (RbEdit.IsChecked == true) mode = "Edit";
            if (RbCopy.IsChecked == true) mode = "Copy";

            // 2. 根据模式执行
            if (mode == "Copy")
            {
                // --- 复制模式 ---
                string defaultPath = Path.GetDirectoryName(i.FullPath);
                if (_activeProject != null && Directory.Exists(_activeProject.Path))
                {
                    defaultPath = _activeProject.Path;
                }

                var d = new CopyMoveDialog(defaultPath, i.Name, false);
                d.Title = "复制副本";

                if (d.ShowDialog() == true)
                {
                    OpenFileSmart(i.FullPath, "Copy", Path.Combine(d.SelectedPath, d.FileName));

                    if (_activeProject != null && d.SelectedPath.StartsWith(_activeProject.Path))
                        RefreshProjectTree();
                }
            }
            else
            {
                // --- 【核心修复】只读/编辑模式 ---
                // 之前这里漏了代码，导致双击没反应
                OpenFileSmart(i.FullPath, mode);
            }
        }
        private void OpenFileSmart(string s, string m, string d = null) { string f = s; if (m == "Copy" && !string.IsNullOrEmpty(d)) { try { File.Copy(s, d, true); f = d; m = "Edit"; } catch { return; } } string x = Path.GetExtension(f).ToLower(); if (x == ".dwg" || x == ".dxf") CadService.OpenDwg(f, m); else Process.Start(new ProcessStartInfo(f) { UseShellExecute = true }); }
        private bool CheckFileFilter(string x, string f)
        {
            if (string.IsNullOrEmpty(f) || f == "所有格式") return true;
            if (f == "DWG图纸" && x.Contains("dwg")) return true;
            if (f == "办公文档" && ".doc.docx.xls.xlsx.ppt.pptx.wps.txt".Contains(x)) return true;
            if (f == "图片" && ".jpg.jpeg.png.bmp.gif".Contains(x)) return true;
            if (f == "PDF" && x == ".pdf") return true;
            if (f == "压缩包" && ".zip.rar.7z".Contains(x)) return true;
            return false;
        }
        private string GetIconForExtension(string x)
        {
            if (x.Contains("dwg")) return "📐";
            if (".doc.docx.xls.xlsx.ppt.pptx.wps.txt".Contains(x)) return "📄";
            if (".jpg.jpeg.png.bmp.gif".Contains(x)) return "🖼️";
            if (x.Contains("pdf")) return "📕";
            if (".zip.rar.7z".Contains(x)) return "📦";
            return "📃";
        }
    }
    // 用于界面箭头旋转的转换器
    public class BooleanToAngleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (bool)value ? 45.0 : 0.0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}