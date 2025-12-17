// 【关键修改：引用新拆分的命名空间】
using CadAtlasManager.Models;
using CadAtlasManager.UI;
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
        public ObservableCollection<FileSystemItem> PlotTreeItems { get; set; }
        public ObservableCollection<ProjectItem> ProjectList { get; set; }

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
            PlotTreeItems = new ObservableCollection<FileSystemItem>();
            ProjectList = new ObservableCollection<ProjectItem>();

            InitializeComponent();

            FileTree.ItemsSource = Items;
            ProjectTree.ItemsSource = ProjectTreeItems;
            PlotTree.ItemsSource = PlotTreeItems;
            CbProjects.ItemsSource = ProjectList;

            LoadConfig();
        }

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

        // =========================================================
        // 【模块 4】图纸工作台逻辑
        // =========================================================
        private void RefreshPlotTree()
        {
            PlotTreeItems.Clear();
            if (_activeProject == null) return;

            // 1. 确保有输出路径配置
            if (string.IsNullOrEmpty(_activeProject.OutputPath))
            {
                _activeProject.OutputPath = Path.Combine(_activeProject.Path, "_Plot");
            }

            // 2. 自动创建 _Plot 文件夹
            if (!Directory.Exists(_activeProject.OutputPath))
            {
                try { Directory.CreateDirectory(_activeProject.OutputPath); } catch { }
            }

            if (!Directory.Exists(_activeProject.OutputPath)) return;

            // 3. 加载根节点
            var root = CreateItem(_activeProject.OutputPath, ExplorerItemType.Folder, true);
            root.Name = "🖨️ 输出归档 (_Plot)";
            root.TypeIcon = "";

            // 4. 加载内容
            LoadPlotSubItems(root);
            PlotTreeItems.Add(root);
        }

        // =================================================================
        // 【重构】批量打印流程 (智能筛选 + 统一参数)
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

                    // 获取当前文件指纹
                    string currentTdUpdate = CadService.GetSmartFingerprint(item.FullPath);

                    // 判断是否过期
                    bool isOutdated = PlotMetaManager.IsOutdated(outputDir, pdfName, dwgName, currentTdUpdate);

                    candidates.Add(new PlotCandidate
                    {
                        FileName = dwgName,
                        FilePath = item.FullPath,
                        IsOutdated = isOutdated,
                        IsSelected = isOutdated, // 默认勾选过期的文件
                        NewTdUpdate = currentTdUpdate,
                        VersionStatus = isOutdated ? "⚠️ 需更新" : "✅ 最新"
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
            var config = dialog.FinalConfig;       // 打印机、纸张、样式等配置
            var filesToPrint = dialog.ConfirmedFiles; // 用户最终勾选的文件

            if (filesToPrint.Count == 0) return;

            // 7. 开始批量打印循环
            int totalSuccess = 0;
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                foreach (var filePath in filesToPrint)
                {
                    // 调用 CadService.BatchPlotByTitleBlocks
                    // 这个方法现在接收 config 对象，并在内部自动处理文档打开/关闭
                    int sheets = CadService.BatchPlotByTitleBlocks(filePath, outputDir, config);

                    if (sheets > 0)
                    {
                        // 打印成功，更新元数据记录
                        var c = candidates.First(x => x.FilePath == filePath);
                        string pdfBaseName = Path.GetFileNameWithoutExtension(c.FileName) + ".pdf";

                        PlotMetaManager.SaveRecord(outputDir, pdfBaseName, c.FileName, c.NewTdUpdate);
                        totalSuccess += sheets;
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
                    ClearAllSelection(PlotTreeItems);

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
        private List<FileSystemItem> GetAllSelectedItems() { var l = new List<FileSystemItem>(); CollectSelected(Items, l); CollectSelected(ProjectTreeItems, l); CollectSelected(PlotTreeItems, l); return l; }
        private void CollectSelected(ObservableCollection<FileSystemItem> s, List<FileSystemItem> r) { if (s == null) return; foreach (var i in s) { if (i.IsItemSelected) r.Add(i); CollectSelected(i.Children, r); } }
        private void SelectRange(ObservableCollection<FileSystemItem> root, FileSystemItem s, FileSystemItem e) { var l = new List<FileSystemItem>(); FlattenTree(root, l); int i1 = l.IndexOf(s), i2 = l.IndexOf(e); if (i1 != -1 && i2 != -1) for (int i = Math.Min(i1, i2); i <= Math.Max(i1, i2); i++) l[i].IsItemSelected = true; }
        private void FlattenTree(ObservableCollection<FileSystemItem> n, List<FileSystemItem> r) { foreach (var node in n) { r.Add(node); if (node.IsExpanded) FlattenTree(node.Children, r); } }
        private TreeViewItem GetTreeViewItemUnderMouse(DependencyObject e) { while (e != null && !(e is TreeViewItem)) e = VisualTreeHelper.GetParent(e); return e as TreeViewItem; }
        private void MenuItem_CopyInPlace_Click_Legacy(object sender, RoutedEventArgs e) { /* 保留旧逻辑引用防止报错，实际使用新的 CopyMove */ }
        private void MenuItem_Remove_Click(object sender, RoutedEventArgs e) { if (FileTree.SelectedItem is FileSystemItem i && i.IsRoot) { _loadedAtlasFolders.Remove(i.FullPath); Items.Remove(i); SaveConfig(); } }
        private void BtnRemoveProject_Click(object sender, RoutedEventArgs e) { if (CbProjects.SelectedItem is ProjectItem p) { ProjectList.Remove(p); ProjectTreeItems.Clear(); PlotTreeItems.Clear(); _activeProject = null; SaveConfig(); } }
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

            // 加载历史记录
            PlotMetaManager.LoadHistory(plotDir);

            // 递归校验所有节点
            Mouse.OverrideCursor = Cursors.Wait; // 显示忙碌光标
            try
            {
                CheckItemsVersion(PlotTreeItems);
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

        // 单个文件校验逻辑
        private void ValidatePdfVersion(FileSystemItem pdfItem)
        {
            string pdfName = pdfItem.Name;
            // 假设命名规则是：[DwgName].pdf
            string dwgName = Path.GetFileNameWithoutExtension(pdfName) + ".dwg";
            string sourceDwgPath = Path.Combine(_activeProject.Path, dwgName);

            // 如果源 DWG 不存在，可能是删了或者改名了
            if (!File.Exists(sourceDwgPath))
            {
                pdfItem.VersionStatus = "❓ 源文件丢失";
                pdfItem.StatusColor = Brushes.Gray;
                return;
            }

            // 获取源文件当前指纹
            string currentTdUpdate = CadService.GetSmartFingerprint(sourceDwgPath);

            // 检查是否过期
            bool isOutdated = PlotMetaManager.IsOutdated(_activeProject.OutputPath, pdfName, dwgName, currentTdUpdate);

            if (isOutdated)
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

        // 2. 右键菜单：打开源 DWG 文件
        private void MenuItem_OpenSourceDwg_Click(object sender, RoutedEventArgs e)
        {
            var item = GetSelectedItem();
            if (item != null && item.Type == ExplorerItemType.File && _activeProject != null)
            {
                // 反推源文件路径
                string dwgName = Path.GetFileNameWithoutExtension(item.Name) + ".dwg";
                string sourceDwgPath = Path.Combine(_activeProject.Path, dwgName);

                if (File.Exists(sourceDwgPath))
                {
                    CadService.OpenDwg(sourceDwgPath, "Edit");
                }
                else
                {
                    MessageBox.Show($"找不到对应的源文件：\n{sourceDwgPath}\n\n可能源文件已重命名或删除。", "错误");
                }
            }
        }
        private void BtnVersion_Click(object sender, RoutedEventArgs e) => MessageBox.Show(_versionInfo);
        private void ProjectTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) { _dragStartPoint = e.GetPosition(null); var t = GetTreeViewItemUnderMouse(e.OriginalSource as DependencyObject); if (t != null) _draggedItem = t.DataContext as FileSystemItem; OnTreeItemClick(sender, e); }
        private void ProjectTree_MouseMove(object sender, MouseEventArgs e) { if (e.LeftButton == MouseButtonState.Pressed && _draggedItem != null) { var diff = _dragStartPoint - e.GetPosition(null); if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance || Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance) { var s = GetAllSelectedItems(); if (!s.Contains(_draggedItem)) { s.Clear(); s.Add(_draggedItem); } DataObject d = new DataObject(DataFormats.FileDrop, s.Select(i => i.FullPath).ToArray()); DragDrop.DoDragDrop(ProjectTree, d, DragDropEffects.Copy | DragDropEffects.Move); _draggedItem = null; } } }
        private void MenuItem_Open_Click(object sender, RoutedEventArgs e) { var i = GetSelectedItem(); if (i != null && i.Type == ExplorerItemType.File) ExecuteFileAction(i); }
        private FileSystemItem GetSelectedItem()
        {
            if (ProjectTree.IsVisible && ProjectTree.SelectedItem != null) return ProjectTree.SelectedItem as FileSystemItem;
            if (PlotTree.IsVisible && PlotTree.SelectedItem != null) return PlotTree.SelectedItem as FileSystemItem;
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
                ClearAllSelection(PlotTreeItems);

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