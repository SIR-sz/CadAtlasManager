// 【关键修改：引用新拆分的命名空间】
using Autodesk.AutoCAD.ApplicationServices;
using CadAtlasManager.Models;
using CadAtlasManager.UI;
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
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
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
        public ObservableCollection<FileSystemItem> ProjectFileListItems { get; set; }
        // 新增：PlotFileListItems 用于右侧列表
        public ObservableCollection<FileSystemItem> PlotFileListItems { get; set; }
        private List<string> _loadedAtlasFolders = new List<string>();
        private ProjectItem _activeProject = null;
        private readonly string _versionInfo =
        "                   CAD图集项目管理系统 v1.0\n" +

        " 专为工程设计打造的图纸全生命周期管理平台\n\n" +
        "【核心功能】\n\n" +
        "① 智能批量打印\n" +
        "   ● 自动识别图框、智能旋转与纸张匹配\n" +
        "   ● 支持模型/布局空间及“自动+手动”混合模式\n\n" +
        "② 项目化图档管理\n" +
        "   ● 资料库 / 项目 / 成果 三级架构\n" +
        "   ● 支持全格式文件管理、一键打包与备注系统\n\n" +
        "③ 版本指纹校验\n" +
        "   ● 实时监控 PDF 与 DWG 内容一致性\n" +
        "   ● 智能标记“过期”图纸，源文件一键反查\n\n" +
        "──────────────────────────\n" +
        "  高效 · 安全 · 智能 — 支持定制拓展开发\n" +
        "──────────────────────────\n\n" +
        "作      者：lsz\n" +
        "联系 QQ ：3956376422\n" +
        "微信公众号：CAD与Office二次开发";
        private FileSystemItem _currentActiveItem = null;
        private FileSystemItem _lastActiveProjectItem = null; // 记录当前变红的文件夹
        // 放在 AtlasView 类内部
        // 1. 扩展状态枚举
        private enum PdfStatus { Latest, Expired, MissingSource, NeedRemerge, Unknown }

        private Point _dragStartPoint;
        private FileSystemItem _draggedItem;

        // 多选与备注相关
        private FileSystemItem _lastSelectedItem = null;
        private FileSystemItem _currentRemarkItem = null; // 当前正在写备注的文件

        // [添加到 AtlasView.xaml.cs 的字段声明区]
        private string _currentProjectFolderPath = ""; // 追踪当前项目视图路径

        private string _currentPlotFolderPath = ""; // 追踪图纸工作台当前显示的文件夹路径

        private List<string> _projectInternalClipboard = new List<string>();// 专门用于项目工作台内部流转的剪贴板

        private FileSystemItem _pendingBindingPdf = null; // 记录当前等待绑定的 PDF 对象

        private readonly List<string> _allowedExtensions = new List<string>
        {
            ".dwg", ".dxf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".wps", ".pdf", ".txt",
            ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".mp4", ".avi", ".mov", //
            ".zip", ".rar", ".7z", ".pat"
        };

        public AtlasView()
        {
            Items = new ObservableCollection<FileSystemItem>();
            ProjectTreeItems = new ObservableCollection<FileSystemItem>();
            PlotFolderItems = new ObservableCollection<FileSystemItem>();
            PlotFileListItems = new ObservableCollection<FileSystemItem>();
            ProjectFileListItems = new ObservableCollection<FileSystemItem>(); // 初始化
            ProjectList = new ObservableCollection<ProjectItem>();

            InitializeComponent();

            FileTree.ItemsSource = Items;
            ProjectTree.ItemsSource = ProjectTreeItems;
            PlotFolderTree.ItemsSource = PlotFolderItems;
            PlotFileList.ItemsSource = PlotFileListItems;
            ProjectFileList.ItemsSource = ProjectFileListItems; // 绑定数据源

            CbProjects.ItemsSource = ProjectList;
            LoadConfig();
        }

        private void ProjectTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var folder = e.NewValue as FileSystemItem;
            LoadProjectFileListItems(folder);
        }
        private void UpdateActiveFolderHighlight(FileSystemItem newItem)
        {
            // 1. 如果之前有激活项，先恢复原状
            if (_currentActiveItem != null)
            {
                _currentActiveItem.IsActive = false;
            }

            // 2. 设置新项为激活状态
            if (newItem != null && newItem.IsDirectory)
            {
                newItem.IsActive = true;
                _currentActiveItem = newItem;
            }
        }
        private void LoadProjectFileListItems(FileSystemItem folder)
        {
            if (folder == null) return;

            // --- 【核心修复：同步左侧树的状态】 ---
            // 1. 在树中找到对应的真实节点对象 (防止是从明细列表传入的新实例)
            FileSystemItem treeNode = FindItemInTree(ProjectTreeItems, folder.FullPath);

            // 2. 状态重置：清除树中所有项的“蓝色底”和“红字”
            // 这解决了“B文件夹蓝底不消失”以及“打包路径不准”的问题
            ClearAllSelection(ProjectTreeItems);
            if (_lastActiveProjectItem != null)
            {
                _lastActiveProjectItem.IsActive = false;
            }

            // 3. 状态激活：如果找到了树节点，同步设置它的状态
            if (treeNode != null)
            {
                treeNode.IsActive = true;       // 变为红字
                treeNode.IsItemSelected = true; // 变为蓝色底 (确保打包逻辑锁定到此路径)
                treeNode.IsExpanded = true;     // 确保父级是展开的
                _lastActiveProjectItem = treeNode;
            }
            // ------------------------------------

            _currentProjectFolderPath = folder.FullPath;
            ProjectFileListItems.Clear();
            if (!Directory.Exists(folder.FullPath)) return;

            try
            {
                string filter = (CbProjectFileType.SelectedItem as ComboBoxItem)?.Content.ToString();

                // 加载子文件夹
                foreach (var dir in Directory.GetDirectories(folder.FullPath))
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(dir);

                    // 1. 原有的逻辑：跳过隐藏文件夹
                    if (dirInfo.Attributes.HasFlag(FileAttributes.Hidden)) continue;

                    // 2. 【新增逻辑】：如果文件夹名称是 "_Plot"，则跳过不显示
                    if (dirInfo.Name.Equals("_Plot", StringComparison.OrdinalIgnoreCase)) continue;

                    var item = CreateItem(dir, ExplorerItemType.Folder);
                    ProjectFileListItems.Add(item);
                }

                // 加载文件
                foreach (var file in Directory.GetFiles(folder.FullPath))
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (_allowedExtensions.Contains(ext) && CheckFileFilter(ext, filter))
                    {
                        var item = CreateItem(file, ExplorerItemType.File);
                        // 【修改位置】：由 GetCreationTime 改为 GetLastWriteTime
                        item.FileDate = File.GetLastWriteTime(file).ToString("yyyy-MM-dd HH:mm");
                        ProjectFileListItems.Add(item);
                    }
                }
            }
            catch { }
        }
        // [修改方法：RefreshPlotTree]

        private void RefreshPlotTree()
        {
            if (_activeProject == null || !Directory.Exists(_activeProject.Path)) return;

            // 1. 记录当前所有展开的路径
            List<string> expandedPaths = new List<string>();
            GetExpandedPaths(PlotFolderItems, expandedPaths);

            // 2. 清空并重建树
            PlotFolderItems.Clear();
            PlotFileListItems.Clear();

            var stageDirs = Directory.GetDirectories(_activeProject.Path, "_Plot", SearchOption.AllDirectories); //

            foreach (var plotPath in stageDirs) //
            {
                string stageDir = Path.GetDirectoryName(plotPath);
                string stageName = (stageDir == _activeProject.Path) ? "项目根目录" : Path.GetFileName(stageDir); //

                var stageNode = new FileSystemItem //
                {
                    Name = stageName,
                    FullPath = stageDir,
                    Type = ExplorerItemType.Folder,
                    TypeIcon = "\uD83C\uDFD7\uFE0F",
                    IsExpanded = true
                };

                var itemSplit = CreateItem(plotPath, ExplorerItemType.Folder); //
                itemSplit.Name = "\uD83D\uDCC4 分项 PDF"; //
                LoadPlotFoldersOnly(itemSplit, "Combined"); //

                string combinedPath = Path.Combine(plotPath, "Combined"); //
                if (!Directory.Exists(combinedPath)) Directory.CreateDirectory(combinedPath);
                var itemCombined = CreateItem(combinedPath, ExplorerItemType.Folder); //
                itemCombined.Name = "\uD83D\uDCD1 成果 PDF"; //

                stageNode.Children.Add(itemSplit);
                stageNode.Children.Add(itemCombined);
                PlotFolderItems.Add(stageNode); //
            }

            // 3. 恢复状态
            RestorePlotTreeState(PlotFolderItems, expandedPaths, _currentPlotFolderPath);
        }
        // [AtlasView.xaml.cs]
        // 处理图纸工作台的快捷键
        // [AtlasView.xaml.cs] - 处理图纸工作台的快捷键
        private void PlotInternal_KeyDown(object sender, KeyEventArgs e)
        {
            // 仅在“图纸工作台”选项卡激活时生效
            if (MainTabControl.SelectedIndex != 2) return;

            if ((e.Key == Key.C || e.Key == Key.X) && Keyboard.Modifiers == ModifierKeys.Control)
            {
                var selected = GetAllSelectedItems();
                if (selected.Count > 0)
                {
                    var paths = new System.Collections.Specialized.StringCollection();
                    foreach (var item in selected) paths.Add(item.FullPath);

                    DataObject data = new DataObject();
                    data.SetFileDropList(paths);

                    if (e.Key == Key.X)
                    {
                        byte[] moveEffect = new byte[] { 2, 0, 0, 0 };
                        MemoryStream dropEffect = new MemoryStream(moveEffect);
                        data.SetData("Preferred DropEffect", dropEffect);
                    }
                    Clipboard.SetDataObject(data);
                }
                e.Handled = true;
            }
            else if (e.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control)
            {
                // 调用你之前新增的图纸粘贴逻辑
                ExecutePlotPaste();
                e.Handled = true;
            }
        }
        // [AtlasView.xaml.cs] - 这是一个完全新增的方法，请直接粘贴到类中
        private void ExecutePlotPaste()
        {
            // 1. 检查系统剪贴板是否有文件
            if (!Clipboard.ContainsFileDropList()) return;
            var filePaths = Clipboard.GetFileDropList();

            // 2. 确定目标目录（使用当前图纸工作台记录的路径）
            string targetDir = _currentPlotFolderPath;
            if (string.IsNullOrEmpty(targetDir) || !Directory.Exists(targetDir)) return;

            // 3. 识别是否为“剪切”操作 (Move 效果)
            bool isMove = false;
            IDataObject data = Clipboard.GetDataObject();
            if (data != null && data.GetDataPresent("Preferred DropEffect"))
            {
                using (MemoryStream ms = (MemoryStream)data.GetData("Preferred DropEffect"))
                {
                    if (ms.ReadByte() == 2) isMove = true;
                }
            }

            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                foreach (string sourcePath in filePaths)
                {
                    if (!File.Exists(sourcePath) && !Directory.Exists(sourcePath)) continue;

                    string fileName = Path.GetFileName(sourcePath);
                    string destPath = Path.Combine(targetDir, fileName);

                    // 自动冲突处理：如果目标已存在，调用你已有的 GenerateInternalUniquePath 重命名
                    if (File.Exists(destPath) || Directory.Exists(destPath))
                    {
                        destPath = GenerateInternalUniquePath(targetDir, fileName);
                    }

                    // 4. 执行物理文件操作
                    if (File.Exists(sourcePath))
                    {
                        if (isMove) File.Move(sourcePath, destPath);
                        else File.Copy(sourcePath, destPath);
                    }
                    else if (Directory.Exists(sourcePath))
                    {
                        if (targetDir.StartsWith(sourcePath, StringComparison.OrdinalIgnoreCase)) continue;

                        if (isMove) Directory.Move(sourcePath, destPath);
                        else CopyProjectDirInternal(sourcePath, destPath); // 调用你已有的递归复制
                    }
                }

                if (isMove) Clipboard.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("图纸粘贴失败: " + ex.Message);
            }
            finally
            {
                Mouse.OverrideCursor = null;
                RefreshPlotTree(); // 刷新界面
            }
        }
        // [AtlasView.xaml.cs] - 确保类中包含此方法
        private void ProjectInternal_KeyDown(object sender, KeyEventArgs e)
        {
            // 仅在“项目工作台”选项卡激活时生效
            if (MainTabControl.SelectedIndex != 1) return;

            // Ctrl + C (复制) 或 Ctrl + X (剪切)
            if ((e.Key == Key.C || e.Key == Key.X) && Keyboard.Modifiers == ModifierKeys.Control)
            {
                var selected = GetAllSelectedItems();
                if (selected.Count > 0)
                {
                    var paths = new System.Collections.Specialized.StringCollection();
                    foreach (var item in selected) paths.Add(item.FullPath);

                    // 构造数据对象
                    DataObject data = new DataObject();
                    data.SetFileDropList(paths);

                    // 如果是剪切，设置 Preferred DropEffect 为 Move
                    if (e.Key == Key.X)
                    {
                        // 0x2 表示 Move (剪切)
                        byte[] moveEffect = new byte[] { 2, 0, 0, 0 };
                        MemoryStream dropEffect = new MemoryStream(moveEffect);
                        data.SetData("Preferred DropEffect", dropEffect);
                    }

                    // 放入系统剪贴板
                    Clipboard.SetDataObject(data);
                }
                e.Handled = true;
            }
            // Ctrl + V (粘贴)
            else if (e.Key == Key.V && Keyboard.Modifiers == ModifierKeys.Control)
            {
                // 调用项目工作台的粘贴逻辑
                ExecuteExternalCompatiblePaste();
                e.Handled = true;
            }
        }
        private void ExecuteExternalCompatiblePaste()
        {
            // 1. 从系统剪贴板获取文件列表
            if (!Clipboard.ContainsFileDropList()) return;
            var filePaths = Clipboard.GetFileDropList();

            // 2. 确定目标目录
            string targetDir = _currentProjectFolderPath;
            if (string.IsNullOrEmpty(targetDir) || !Directory.Exists(targetDir))
            {
                if (_activeProject != null) targetDir = _activeProject.Path;
                else return;
            }

            // 3. 检查当前剪贴板是否是“剪切”操作
            bool isMove = false;
            IDataObject data = Clipboard.GetDataObject();
            if (data.GetDataPresent("Preferred DropEffect"))
            {
                using (MemoryStream ms = (MemoryStream)data.GetData("Preferred DropEffect"))
                {
                    if (ms.ReadByte() == 2) isMove = true;
                }
            }

            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                foreach (string sourcePath in filePaths)
                {
                    if (!File.Exists(sourcePath) && !Directory.Exists(sourcePath)) continue;

                    string fileName = Path.GetFileName(sourcePath);
                    string destPath = Path.Combine(targetDir, fileName);

                    // 冲突检查：同名则自动重命名
                    if (File.Exists(destPath) || Directory.Exists(destPath))
                    {
                        destPath = GenerateInternalUniquePath(targetDir, fileName);
                    }

                    // 4. 执行物理操作
                    if (File.Exists(sourcePath))
                    {
                        if (isMove) File.Move(sourcePath, destPath); // 剪切：移动文件
                        else File.Copy(sourcePath, destPath);        // 复制：拷贝文件

                        // 同步备注
                        string remark = RemarkManager.GetRemark(sourcePath);
                        if (!string.IsNullOrEmpty(remark)) RemarkManager.SaveRemark(destPath, remark);
                    }
                    else if (Directory.Exists(sourcePath))
                    {
                        if (targetDir.StartsWith(sourcePath, StringComparison.OrdinalIgnoreCase)) continue;

                        if (isMove) Directory.Move(sourcePath, destPath);
                        else CopyProjectDirInternal(sourcePath, destPath);
                    }
                }

                // 5. 如果是剪切，粘贴完后清空剪贴板（防止重复移动报错）
                if (isMove) Clipboard.Clear();
            }
            catch (Exception ex) { MessageBox.Show("粘贴操作失败: " + ex.Message); }
            finally
            {
                Mouse.OverrideCursor = null;
                RefreshProjectTree();
            }
        }
        private void ExecuteInternalProjectPaste()
        {
            if (_projectInternalClipboard == null || _projectInternalClipboard.Count == 0) return;

            // 确定目标目录：当前选中的文件夹路径或项目根目录
            string targetDir = _currentProjectFolderPath;
            if (string.IsNullOrEmpty(targetDir) || !Directory.Exists(targetDir))
            {
                if (_activeProject != null) targetDir = _activeProject.Path;
                else return;
            }

            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                foreach (string sourcePath in _projectInternalClipboard)
                {
                    if (!File.Exists(sourcePath) && !Directory.Exists(sourcePath)) continue;

                    string fileName = Path.GetFileName(sourcePath);
                    string destPath = Path.Combine(targetDir, fileName);

                    // 自动重命名逻辑：如果目标已存在，生成“ - 副本”后缀
                    if (File.Exists(destPath) || Directory.Exists(destPath))
                    {
                        destPath = GenerateInternalUniquePath(targetDir, fileName);
                    }

                    // 执行物理复制并同步备注
                    if (File.Exists(sourcePath))
                    {
                        File.Copy(sourcePath, destPath);
                        string remark = RemarkManager.GetRemark(sourcePath);
                        if (!string.IsNullOrEmpty(remark)) RemarkManager.SaveRemark(destPath, remark);
                    }
                    else if (Directory.Exists(sourcePath))
                    {
                        // 风险排查：禁止将父文件夹粘贴进自己的子文件夹
                        if (targetDir.StartsWith(sourcePath, StringComparison.OrdinalIgnoreCase)) continue;
                        CopyProjectDirInternal(sourcePath, destPath);
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("项目内粘贴失败: " + ex.Message); }
            finally
            {
                Mouse.OverrideCursor = null;
                RefreshProjectTree(); // 刷新界面
            }
        }

        // 内部使用的唯一路径生成器
        private string GenerateInternalUniquePath(string folder, string fileName)
        {
            string nameOnly = Path.GetFileNameWithoutExtension(fileName);
            string ext = Path.GetExtension(fileName);
            int count = 1;
            string newPath = Path.Combine(folder, $"{nameOnly} - 副本{ext}");

            while (File.Exists(newPath) || Directory.Exists(newPath))
            {
                count++;
                newPath = Path.Combine(folder, $"{nameOnly} - 副本({count}){ext}");
            }
            return newPath;
        }

        // 递归复制文件夹及其备注
        private void CopyProjectDirInternal(string source, string dest)
        {
            Directory.CreateDirectory(dest);
            foreach (var file in Directory.GetFiles(source))
            {
                string dFile = Path.Combine(dest, Path.GetFileName(file));
                File.Copy(file, dFile);
                string rem = RemarkManager.GetRemark(file);
                if (!string.IsNullOrEmpty(rem)) RemarkManager.SaveRemark(dFile, rem);
            }
            foreach (var sub in Directory.GetDirectories(source))
                CopyProjectDirInternal(sub, Path.Combine(dest, Path.GetFileName(sub)));
        }
        // 专门为图纸树定制的恢复逻辑
        private void RestorePlotTreeState(ObservableCollection<FileSystemItem> nodes, List<string> expandedPaths, string targetPath)
        {
            foreach (var node in nodes)
            {
                if (expandedPaths.Contains(node.FullPath)) node.IsExpanded = true;

                if (node.FullPath.Equals(targetPath, StringComparison.OrdinalIgnoreCase))
                {
                    node.IsItemSelected = true;
                    LoadPlotFilesList(node); // 强制刷新右侧 PDF 列表
                }

                if (node.Children.Count > 0)
                    RestorePlotTreeState(node.Children, expandedPaths, targetPath);
            }
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

                    // \u2705 新增过滤：如果文件夹名等于我们要排除的名字（比如 Combined），则跳过
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

        // [修改方法: LoadPlotFilesList]
        private void LoadPlotFilesList(FileSystemItem folder)
        {
            if (folder == null) return;
            // ================== 核心修复开始 ==================
            // 1. 在树的数据源中找到对应的真实节点对象
            // (虽然入参 folder 可能就是那个节点，但为了保险起见，尤其是从外部调用时，重新查找最稳妥)
            FileSystemItem treeNode = FindItemInTree(PlotFolderItems, folder.FullPath);

            // 2. 清除树中所有已高亮的项 (解决“旧的高亮不消失”的问题)
            ClearAllSelection(PlotFolderItems);

            // 3. 激活当前节点 (解决“点击新文件夹不高亮”的问题)
            if (treeNode != null)
            {
                treeNode.IsItemSelected = true;
                // 顺便确保父级展开，防止从代码跳转过来时没展开
                // (注意：这里不要强制 treeNode.IsExpanded = true，否则点击父文件夹想折叠时会自动弹开)
            }
            // ================== 核心修复结束 ==================

            _currentPlotFolderPath = folder.FullPath; // 关键：记录当前图纸查看路径

            PlotFileListItems.Clear(); //
            if (!Directory.Exists(folder.FullPath)) return;

            try
            {
                foreach (var file in Directory.GetFiles(folder.FullPath)) //
                {
                    string ext = Path.GetExtension(file).ToLower();
                    if (".pdf.jpg.jpeg.png.plt.tif.tiff.txt".Contains(ext)) //
                    {
                        var item = CreateItem(file, ExplorerItemType.File);
                        // 【修改位置】：由 GetCreationTime 改为 GetLastWriteTime
                        item.FileDate = File.GetLastWriteTime(file).ToString("yyyy-MM-dd HH:mm");
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
        // 项目工作台 - 文件明细列表全选/反选逻辑
        private void OnSelectAllProjectFiles_Click(object sender, RoutedEventArgs e)
        {
            var cb = sender as CheckBox;
            if (cb == null || ProjectFileListItems == null) return;

            bool isChecked = cb.IsChecked == true;

            // 遍历当前项目列表中的所有项，统一修改选中状态
            foreach (var item in ProjectFileListItems)
            {
                item.IsChecked = isChecked;
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
        // 项目工作台的文件列表双击：始终以 Edit 模式打开
        // [修改文件: UI/AtlasView.xaml.cs]
        private void ProjectFileList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ProjectFileList.SelectedItem is FileSystemItem item)
            {
                if (item.Type == ExplorerItemType.File)
                {
                    OpenFileSmart(item.FullPath, "Edit");
                }
                else if (item.Type == ExplorerItemType.Folder)
                {
                    // 这里会自动触发上面修改后的同步逻辑，让文件夹变红
                    LoadProjectFileListItems(item);
                }
            }
        }
        // [添加到 AtlasView.xaml.cs]
        private bool SyncTreeSelection(ObservableCollection<FileSystemItem> nodes, string targetPath)
        {
            foreach (var node in nodes)
            {
                if (node.FullPath.Equals(targetPath, StringComparison.OrdinalIgnoreCase))
                {
                    // 找到目标节点：清除旧选中，设置新选中并展开
                    ClearAllSelection(ProjectTreeItems); // 使用现有的清除方法
                    node.IsItemSelected = true;
                    node.IsExpanded = true;
                    return true;
                }

                if (node.Children.Count > 0)
                {
                    if (SyncTreeSelection(node.Children, targetPath))
                    {
                        // 如果在子项中找到了，父项也需要展开
                        node.IsExpanded = true;
                        return true;
                    }
                }
            }
            return false;
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
                    if (System.IO.File.Exists(item.FullPath))
                    {
                        // 使用全路径引用，避免 SearchOption 冲突
                        Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(
                            item.FullPath,
                            Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                            Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);

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

            // 刷新逻辑
            var currentFolder = PlotFolderTree.SelectedItem as FileSystemItem;
            if (currentFolder != null)
            {
                LoadPlotFilesList(currentFolder);
            }
            else
            {
                RefreshPlotTree();
            }
            // 双保险：如果通过树刷新没定位到，手动寻找并重载
            if (!string.IsNullOrEmpty(_currentPlotFolderPath))
            {
                var currentItem = FindItemInTree(PlotFolderItems, _currentPlotFolderPath);
                if (currentItem != null)
                {
                    LoadPlotFilesList(currentItem);
                }
            }
        }

        // [修改方法: BtnMergePdf_Click]
        private void BtnMergePdf_Click(object sender, RoutedEventArgs e)
        {
            var targets = PlotFileListItems
                .Where(i => i.IsChecked && i.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                .OrderBy(i => i.Name).ToList();

            if (targets.Count < 2) { MessageBox.Show("请至少勾选 2 个文件。"); return; }

            // 【关键修改】动态获取当前阶段的 _Plot 目录
            string currentPlotDir = GetCurrentPlotDir();
            if (string.IsNullOrEmpty(currentPlotDir)) return;

            string defaultName = Path.GetFileNameWithoutExtension(targets[0].Name) + "_合并";
            var dlg = new RenameDialog(defaultName) { Title = "合并 PDF" };
            if (dlg.ShowDialog() != true) return;

            string saveDir = Path.Combine(currentPlotDir, "Combined"); // 保存到当前阶段的 Combined
            if (!Directory.Exists(saveDir)) Directory.CreateDirectory(saveDir);

            string savePath = Path.Combine(saveDir, dlg.NewName + ".pdf");

            Mouse.OverrideCursor = Cursors.Wait;
            try
            {
                MergePdfFiles(targets.Select(t => t.FullPath).ToList(), savePath);

                // 记录合并关系到当前阶段的元数据中
                PlotMetaManager.SaveCombinedRecord(currentPlotDir, Path.GetFileName(savePath), targets);

                MessageBox.Show("成功合并到当前阶段成果库。");

                // 局部刷新：重新加载当前选中的文件夹内容
                if (PlotFolderTree.SelectedItem is FileSystemItem currentFolder)
                    LoadPlotFilesList(currentFolder);
            }
            catch (Exception ex) { MessageBox.Show("合并失败: " + ex.Message); }
            finally { Mouse.OverrideCursor = null; }
        }
        // =================================================================
        // 【新增功能】打开分项 PDF 对应的源 DWG 文件
        // =================================================================
        private void BtnOpenSourceDwg_Click(object sender, RoutedEventArgs e)
        {
            // 1. 基础检查
            if (_activeProject == null) return;

            // 获取列表选中项 (注意：这里只支持单选打开，如果多选了，只打开第一个或提示)
            var item = PlotFileList.SelectedItem as FileSystemItem;

            if (item == null || item.Type != ExplorerItemType.File || !item.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("请先在列表中选择一个 PDF 文件。", "提示");
                return;
            }

            // 2. 检查是否为“合并文件” (合并文件没有单一的源 DWG，不能打开)
            string plotDir = _activeProject.OutputPath;
            if (PlotMetaManager.IsCombinedFile(plotDir, item.Name))
            {
                MessageBox.Show("这是一个合并后的成果 PDF，包含多个源文件。\n无法直接定位到单一的 DWG 图纸。", "提示");
                return;
            }

            // 3. 开始寻找源文件
            // A. 优先从历史记录反查 (最准确，能处理改名情况)
            string realDwgName = PlotMetaManager.GetSourceDwgName(plotDir, item.Name);

            // B. 历史记录查不到，尝试智能猜测 (去除 _1, _2 等布局后缀)
            // 例如：Drawing1_1.pdf -> 猜测源文件是 Drawing1.dwg
            if (string.IsNullOrEmpty(realDwgName))
            {
                string baseName = System.Text.RegularExpressions.Regex.Replace(Path.GetFileNameWithoutExtension(item.Name), @"_\d+$", "");
                realDwgName = baseName + ".dwg";
            }

            // 4. 定位文件路径
            // 先尝试在项目根目录找
            string sourceDwgPath = Path.Combine(_activeProject.Path, realDwgName);

            if (!File.Exists(sourceDwgPath))
            {
                // 如果根目录没有，可能是文件被整理到子文件夹了
                // 进行全项目递归搜索 (只搜文件名)
                try
                {
                    string onlyFileName = Path.GetFileName(realDwgName);
                    var foundFiles = Directory.GetFiles(_activeProject.Path, onlyFileName, SearchOption.AllDirectories);

                    if (foundFiles.Length > 0)
                    {
                        // 找到了！取第一个匹配项
                        sourceDwgPath = foundFiles[0];
                    }
                    else
                    {
                        // 彻底没找到
                        MessageBox.Show($"未找到源文件：\n{realDwgName}\n\n可能原因：\n1. DWG文件已被重命名或删除\n2. 该PDF不是由当前项目图纸生成的", "文件丢失");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("搜索源文件时出错: " + ex.Message);
                    return;
                }
            }

            // 5. 执行打开
            try
            {
                // "Edit" 模式表示以读写方式打开，方便用户修改
                CadService.OpenDwg(sourceDwgPath, "Edit");
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开 CAD 图纸失败: " + ex.Message);
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

        // [修改方法: RefreshProjectTree]
        // [修改方法: RefreshProjectTree]
        private void RefreshProjectTree()
        {
            if (_activeProject == null || !Directory.Exists(_activeProject.Path)) return;

            // 1. 记录当前所有展开的路径
            List<string> expandedPaths = new List<string>();
            GetExpandedPaths(ProjectTreeItems, expandedPaths);

            // 2. 彻底清空并重建树
            ProjectTreeItems.Clear();
            RemarkManager.LoadRemarks(_activeProject.Path);

            var root = CreateItem(_activeProject.Path, ExplorerItemType.Folder, true);
            root.Name = _activeProject.Name;
            root.TypeIcon = "\uD83C\uDFD7\uFE0F";

            LoadProjectSubItems(root);
            ProjectTreeItems.Add(root);

            // 3. 恢复展开状态，并根据记录的路径重新定位选中项
            RestoreProjectTreeState(ProjectTreeItems, expandedPaths, _currentProjectFolderPath);
        }

        // 专门为项目树定制的恢复逻辑
        private void RestoreProjectTreeState(ObservableCollection<FileSystemItem> nodes, List<string> expandedPaths, string targetPath)
        {
            foreach (var node in nodes)
            {
                if (expandedPaths.Contains(node.FullPath)) node.IsExpanded = true;

                if (node.FullPath.Equals(targetPath, StringComparison.OrdinalIgnoreCase))
                {
                    node.IsItemSelected = true;
                    // 关键：强制刷新右侧列表，确保删除后的文件消失
                    LoadProjectFileListItems(node);
                }

                if (node.Children.Count > 0)
                    RestoreProjectTreeState(node.Children, expandedPaths, targetPath);
            }
        }

        private void LoadProjectSubItems(FileSystemItem parent)
        {
            try
            {
                RemarkManager.LoadRemarks(parent.FullPath);
                foreach (var dir in Directory.GetDirectories(parent.FullPath))
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(dir);
                    if (dirInfo.Attributes.HasFlag(FileAttributes.Hidden)) continue;

                    // 【新增】：同步在左侧树中隐藏 _Plot 文件夹
                    if (dirInfo.Name.Equals("_Plot", StringComparison.OrdinalIgnoreCase)) continue;

                    var sub = CreateItem(dir, ExplorerItemType.Folder);
                    LoadProjectSubItems(sub);
                    parent.Children.Add(sub);
                }
            }
            catch { }
        }


        // [修改方法：MenuItem_BatchPlot_Click]
        // [AtlasView.xaml.cs]
        // [AtlasView.xaml.cs] 完整的方法实现
        private void MenuItem_BatchPlot_Click(object sender, RoutedEventArgs e)
        {
            // 获取选中的 DWG 文件列表
            var selectedDwgs = GetAllSelectedItems()
        .Where(i => i.Type == ExplorerItemType.File && i.FullPath.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
        .ToList();

            if (selectedDwgs.Count == 0) return;

            var candidates = selectedDwgs.Select(d => new PlotCandidate
            {
                FilePath = d.FullPath,
                FileName = System.IO.Path.GetFileName(d.FullPath)
            }).ToList();

            // 弹出对话框
            var dialog = new BatchPlotDialog(candidates);

            // 【优化点 2】：利用对话框阻塞特性，关闭后立即执行刷新
            dialog.ShowDialog();

            // 刷新图纸树（重新扫描磁盘 _Plot 目录）
            RefreshPlotTree();

            // 如果当前选中的就是图纸工作台，强制触发一次右侧列表重载
            if (!string.IsNullOrEmpty(_currentPlotFolderPath))
            {
                var currentItem = FindItemInTree(PlotFolderItems, _currentPlotFolderPath);
                if (currentItem != null)
                {
                    LoadPlotFilesList(currentItem);
                }
            }
        }

        private void ExecuteHybridPlotWorkflow(List<FileSystemItem> dwgs, BatchPlotConfig config)
        {
            var results = new List<PlotFileResult>();
            System.Windows.Input.Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;

            try
            {
                // 第一步：全自动批量执行
                foreach (var dwg in dwgs)
                {
                    string targetDir = Path.Combine(Path.GetDirectoryName(dwg.FullPath), "_Plot");
                    if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);

                    var res = CadService.EnhancedBatchPlot(dwg.FullPath, targetDir, config);
                    results.Add(res);
                }
                System.Windows.Input.Mouse.OverrideCursor = null;

                // 第二步：汇总展示
                int failCount = results.Count(r => !r.IsSuccess);
                if (failCount > 0)
                {
                    string msg = $"自动打印完成。\n成功: {results.Count - failCount} 个\n识别失败: {failCount} 个\n\n是否立即开始手动拾取失败的图纸？";
                    if (MessageBox.Show(msg, "打印确认", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    {
                        // 第三步：手动拾取阶段
                        foreach (var failItem in results.Where(r => !r.IsSuccess))
                        {
                            // 使用 AcApp 引用 AutoCAD 的 Application
                            var acDoc = AcApp.DocumentManager.Open(failItem.FilePath, false);
                            string plotDir = Path.Combine(Path.GetDirectoryName(failItem.FilePath), "_Plot");

                            MessageBox.Show($"正在手动拾取文件：\n{failItem.FileName}\n\n请查看 CAD 命令行选择【手动框选】或【选择图框块】模式进行操作。");
                            CadService.ManualPickAndPlot(acDoc, plotDir, config);

                            acDoc.CloseAndDiscard();
                        }
                    }
                }
                else
                {
                    MessageBox.Show("全自动打印已全部完成！");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show("混合打印流程出错: " + ex.Message);
            }
            finally
            {
                System.Windows.Input.Mouse.OverrideCursor = null;
                RefreshPlotTree();
            }
        }

        private void ExecuteManualPhase(List<PlotFileResult> fails, BatchPlotConfig config)
        {
            foreach (var item in fails)
            {
                try
                {
                    // 打开文件
                    var doc = AcApp.DocumentManager.Open(item.FilePath, false);
                    string targetDir = Path.Combine(Path.GetDirectoryName(item.FilePath), "_Plot");

                    MessageBox.Show($"正在处理：{item.FileName}\n\n请在 CAD 窗口中连续框选区域，按回车结束该文件。");

                    // 调用手动拾取服务
                    int count = CadService.ManualPickAndPlot(doc, targetDir, config);

                    // 处理元数据关联（如果有的话，复用原来的 PlotMetaManager）
                    // ... 

                    doc.CloseAndDiscard(); // 及时关闭，保持 CAD 运行稳定
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"手动打印 [{item.FileName}] 出错: {ex.Message}");
                }
            }
        }

        // [AtlasView.xaml.cs 或对应的逻辑文件]
        private void BtnBatchPlotFromPlot_Click(object sender, RoutedEventArgs e)
        {
            if (_activeProject == null) return;

            // 1. 获取图纸列表中勾选或选中的 PDF 项
            var selectedPdfs = PlotFileListItems.Where(i => i.IsChecked || i.IsItemSelected)
                .Where(i => i.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (selectedPdfs.Count == 0)
            {
                MessageBox.Show("请先在列表中勾选或选择需要重印的 PDF 文件。", "提示");
                return;
            }

            // 2. 转换：寻找每个 PDF 对应的源 DWG 文件对象
            var dwgItemsForPlot = new List<FileSystemItem>();
            var missingDwgs = new List<string>();

            foreach (var pdf in selectedPdfs)
            {
                // 排除合并文件（合并文件无法直接重印，需要更新分项后再合并）
                string plotDir = GetCurrentPlotDir();
                if (PlotMetaManager.IsCombinedFile(plotDir, pdf.Name)) continue;

                // 查找源 DWG 路径
                string realDwgName = PlotMetaManager.GetSourceDwgName(plotDir, pdf.Name);
                if (string.IsNullOrEmpty(realDwgName))
                {
                    string baseName = System.Text.RegularExpressions.Regex.Replace(Path.GetFileNameWithoutExtension(pdf.Name), @"_\d+$", "");
                    realDwgName = baseName + ".dwg";
                }

                string sourceFolder = Path.GetDirectoryName(plotDir);
                string dwgPath = Path.Combine(sourceFolder, realDwgName);

                if (!string.IsNullOrEmpty(dwgPath) && File.Exists(dwgPath))
                {
                    if (!dwgItemsForPlot.Any(i => i.FullPath.Equals(dwgPath, StringComparison.OrdinalIgnoreCase)))
                    {
                        dwgItemsForPlot.Add(CreateItem(dwgPath, ExplorerItemType.File));
                    }
                }
                else
                {
                    missingDwgs.Add(pdf.Name);
                }
            }

            if (dwgItemsForPlot.Count == 0)
            {
                MessageBox.Show("未找到选中 PDF 对应的源 DWG 文件，无法重印。", "提示");
                return;
            }

            // 3. 调起打印对话框
            var candidates = PrepareCandidates(dwgItemsForPlot);

            // --- 核心修改部分 ---
            var dialog = new BatchPlotDialog(candidates)
            {
                Title = "图纸重印 - 批量打印设置 (重印模式)", // 修改标题增加识别度
                IsReprintMode = true                           // 开启重印模式，触发生成“打印目录重印部分.txt”
            };

            if (dialog.ShowDialog() != true) return;

            // 4. 执行后续处理 (如果你的 ExecuteBatchPlotProcess 内部也会涉及生成目录，
            // 请确保该方法也能感知到 IsReprintMode，但根据我们之前的修改，
            // 目录生成逻辑已经集成在对话框内部的“开始打印”按钮中了)
            ExecuteBatchPlotProcess(dwgItemsForPlot, dialog.FinalConfig, dialog.ConfirmedFiles);
        }

        // 提取出的通用打印执行流程，方便复用
        private void ExecuteBatchPlotProcess(List<FileSystemItem> dwgFiles, BatchPlotConfig config, List<string> confirmedPaths)
        {
            int totalSuccess = 0;
            Mouse.OverrideCursor = Cursors.Wait;

            try
            {
                var groupedFiles = dwgFiles.GroupBy(f => Path.GetDirectoryName(f.FullPath));

                foreach (var group in groupedFiles)
                {
                    string sourceFolder = group.Key;
                    string targetPlotDir = Path.Combine(sourceFolder, "_Plot");
                    if (!Directory.Exists(targetPlotDir)) Directory.CreateDirectory(targetPlotDir);

                    PlotMetaManager.LoadHistory(targetPlotDir);

                    foreach (var dwg in group)
                    {
                        if (!confirmedPaths.Contains(dwg.FullPath)) continue;

                        // 执行打印
                        List<string> generatedPdfs = CadService.BatchPlotByTitleBlocks(dwg.FullPath, targetPlotDir, config);

                        if (generatedPdfs.Count > 0)
                        {
                            string freshFingerprint = CadService.GetContentFingerprint(dwg.FullPath);
                            string freshTime = CadService.GetFileTimestamp(dwg.FullPath);
                            foreach (var pdfName in generatedPdfs)
                            {
                                PlotMetaManager.SaveRecord(targetPlotDir, pdfName, dwg.Name, freshFingerprint, freshTime);
                            }
                            totalSuccess += generatedPdfs.Count;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("重印过程中出错: " + ex.Message);
            }
            finally
            {
                Mouse.OverrideCursor = null;
                RefreshPlotTree(); // 刷新归档树以显示最新状态
                MessageBox.Show($"打印完成，共更新 {totalSuccess} 页图纸。");
            }
        }
        // 辅助方法：准备候选列表（保留版本校验核心）
        private List<PlotCandidate> PrepareCandidates(List<FileSystemItem> dwgs)
        {
            var list = new List<PlotCandidate>();
            foreach (var item in dwgs)
            {
                string plotDir = Path.Combine(Path.GetDirectoryName(item.FullPath), "_Plot");
                PlotMetaManager.LoadHistory(plotDir); // 加载局部的历史记录

                string pdfName = Path.GetFileNameWithoutExtension(item.Name) + ".pdf";
                string currentTimestamp = CadService.GetFileTimestamp(item.FullPath);

                // 调用原有 PlotMetaManager 校验逻辑
                bool isLatest = PlotMetaManager.CheckStatus(plotDir, pdfName, item.Name, currentTimestamp,
                    () => CadService.GetContentFingerprint(item.FullPath));

                list.Add(new PlotCandidate
                {
                    FileName = item.Name,
                    FilePath = item.FullPath,
                    IsOutdated = !isLatest,
                    IsSelected = !isLatest,
                    VersionStatus = !isLatest ? "\u26A0\uFE0F 需更新" : "\u2705 最新"
                });
            }
            return list;
        }
        // 3. 更新 UI 显示逻辑
        // 2. 修改 UpdateVersionUi 支持文字替换
        // [AtlasView.xaml.cs]
        private void UpdateVersionUi(FileSystemItem item, PdfStatus status, bool isExternal = false)
        {
            // 如果是外部绑定，设置前缀文字
            string prefix = isExternal ? "外部绑定" : "";

            switch (status)
            {
                case PdfStatus.Latest:
                    // 结果示例：\u2705 最新 或 \u2705 外部绑定最新
                    item.VersionStatus = $"\u2705 {prefix}最新";
                    item.StatusColor = Brushes.Green;
                    break;
                case PdfStatus.Expired:
                    // 结果示例：\u26A0\uFE0F 需更新 或 \u26A0\uFE0F 外部绑定需更新
                    item.VersionStatus = $"\u26A0\uFE0F {prefix}需更新";
                    item.StatusColor = Brushes.Red;
                    break;
                case PdfStatus.NeedRemerge:
                    item.VersionStatus = "\uD83D\uDD04 需重并";
                    item.StatusColor = Brushes.Orange;
                    break;
                case PdfStatus.MissingSource:
                    item.VersionStatus = "\u2753 源缺失";
                    item.StatusColor = Brushes.Gray;
                    break;
                default:
                    item.VersionStatus = "";
                    break;
            }
        }
        // [添加到 AtlasView.xaml.cs]
        private string FindDwgInProject(string basePath, string fileName)
        {
            if (string.IsNullOrEmpty(basePath) || string.IsNullOrEmpty(fileName)) return null;

            try
            {
                // 确保文件名不含路径
                string onlyName = Path.GetFileName(fileName);

                // 在项目根目录下执行递归搜索
                var foundFiles = Directory.GetFiles(basePath, onlyName, SearchOption.AllDirectories);

                // 如果找到了，返回第一个匹配项的完整路径
                return foundFiles.Length > 0 ? foundFiles[0] : null;
            }
            catch (Exception ex)
            {
                // 记录错误日志，防止搜索过程中因权限等问题崩溃
                System.Diagnostics.Debug.WriteLine("搜索 DWG 文件出错: " + ex.Message);
                return null;
            }
        }
        private void LoadPlotSubItems(FileSystemItem parent)
        {
            try
            {
                RemarkManager.LoadRemarks(parent.FullPath);
                // 仅加载 PDF, JPG, PNG, PLT，.tif，tiff
                string plotExts = ".pdf.jpg.jpeg.png.plt.tif.tiff";

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
        // 在 AtlasView.xaml.cs 中添加
        // [AtlasView.xaml.cs]
        private void MenuItem_InsertXref_Click(object sender, RoutedEventArgs e)
        {
            // 1. 获取当前选中的文件
            var item = GetSelectedItem();

            if (item != null && item.Type == ExplorerItemType.File)
            {
                string fullPath = item.FullPath;
                string ext = System.IO.Path.GetExtension(fullPath).ToLower();

                // 2. 将文件路径复制到系统剪贴板，方便用户在弹出窗口中 Ctrl+V
                System.Windows.Clipboard.SetText(fullPath);

                // 3. 根据后缀名动态选择对应的 CAD 附着命令
                if (ext == ".dwg" || ext == ".dxf")
                {
                    // DWG 图纸使用外部参照 (XATTACH)
                    CadService.OpenXrefDialog();
                }
                else if (".jpg.jpeg.png.bmp.gif.tif.tiff".Contains(ext))
                {
                    // 图片使用图像附着 (IMAGEATTACH)
                    CadService.OpenImageAttachDialog();
                }
                else if (ext == ".pdf")
                {
                    // PDF 文件使用 PDF 参考底图附着 (PDFATTACH)
                    CadService.OpenPdfAttachDialog();
                }
                else
                {
                    System.Windows.MessageBox.Show("该文件格式不支持作为参照插入。");
                }
            }
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
                TypeIcon = type == ExplorerItemType.Folder ? "\uD83D\uDCC1" : GetIconForExtension(ext),
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
                        if (item.Type == ExplorerItemType.File)
                        {
                            if (System.IO.File.Exists(item.FullPath))
                            {
                                Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(
                                    item.FullPath,
                                    Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                                    Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                            }
                        }
                        else
                        {
                            if (System.IO.Directory.Exists(item.FullPath))
                            {
                                Microsoft.VisualBasic.FileIO.FileSystem.DeleteDirectory(
                                    item.FullPath,
                                    Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                                    Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                            }
                        }

                        // 同步删除备注
                        RemarkManager.HandleDelete(item.FullPath);
                    }
                    RefreshProjectTree(); RefreshPlotTree();
                    if (!string.IsNullOrEmpty(_currentProjectFolderPath))
                    {
                        // 尝试在重新生成后的树中寻找之前的文件夹对象
                        var currentItem = FindItemInTree(ProjectTreeItems, _currentProjectFolderPath);
                        if (currentItem != null)
                        {
                            // 显式触发一次右侧列表加载，确保磁盘上的删除结果被刷新到 UI
                            LoadProjectFileListItems(currentItem);
                        }
                    }
                }
                catch (System.Exception ex) { MessageBox.Show("删除失败: " + ex.Message); }
            }
        }
        // [添加到 AtlasView.xaml.cs]
        private FileSystemItem FindItemInTree(ObservableCollection<FileSystemItem> nodes, string path)
        {
            foreach (var node in nodes)
            {
                if (node.FullPath.Equals(path, StringComparison.OrdinalIgnoreCase))
                    return node;

                if (node.Children.Count > 0)
                {
                    var found = FindItemInTree(node.Children, path);
                    if (found != null) return found;
                }
            }
            return null;
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
                    // 1. 清除所有左侧树状结构的选中状态
                    ClearAllSelection(Items);
                    ClearAllSelection(ProjectTreeItems);
                    ClearAllSelection(PlotFolderItems);

                    // 2. 【核心修复】清除“项目工作台”右侧列表的勾选 (新增)
                    if (ProjectFileListItems != null)
                    {
                        foreach (var item in ProjectFileListItems) item.IsChecked = false;
                    }

                    // 3. 【核心修复】清除“图纸工作台”右侧列表的勾选 (保留并增加 null 检查)
                    if (PlotFileListItems != null)
                    {
                        foreach (var item in PlotFileListItems) item.IsChecked = false;
                    }

                    clickedItem.IsItemSelected = true;
                    _lastSelectedItem = clickedItem;
                }
            }
        }



        // =================================================================
        // 【修改】一键打包功能 (使用自定义 ZipSaveDialog)
        // =================================================================
        private void MenuItem_Zip_Click(object sender, RoutedEventArgs e)
        {
            // 1. 获取选中项：优先取勾选的，没勾选取选中的
            var selectedItems = GetAllSelectedItems();
            if (selectedItems.Count == 0)
            {
                var current = GetSelectedItem();
                if (current != null) selectedItems.Add(current);
            }

            if (selectedItems.Count == 0) return;

            // 2. 确定保存路径
            string firstPath = selectedItems[0].FullPath;

            // 【关键修改】：不再判断是否为文件夹，统一取父目录 Path.GetDirectoryName
            // 这样如果打包文件夹 A，zip 会默认出现在 A 的旁边，而不是 A 的里面
            string defaultPath = Path.GetDirectoryName(firstPath);

            // 防御性处理：如果是根目录，则维持原样
            if (string.IsNullOrEmpty(defaultPath)) defaultPath = firstPath;

            string defaultName = $"打包_{DateTime.Now:yyyyMMdd_HHmm}.zip";
            var dlg = new ZipSaveDialog(defaultPath, defaultName);

            if (dlg.ShowDialog() == true)
            {
                string zipFullPath = Path.Combine(dlg.SelectedPath, dlg.FileName);
                try
                {
                    // 如果目标文件已存在，先尝试删除（防止占用报错）
                    if (File.Exists(zipFullPath))
                    {
                        if (MessageBox.Show("压缩包已存在，是否覆盖？", "提示", MessageBoxButton.YesNo) == MessageBoxResult.No)
                            return;
                        File.Delete(zipFullPath);
                    }

                    using (ZipArchive archive = ZipFile.Open(zipFullPath, ZipArchiveMode.Create))
                    {
                        foreach (var item in selectedItems)
                        {
                            if (item.Type == ExplorerItemType.File)
                                archive.CreateEntryFromFile(item.FullPath, item.Name);
                            else
                                AddFolderToZip(archive, item.FullPath, item.Name);
                        }
                    }

                    if (MessageBox.Show("打包成功！是否打开所在文件夹？", "完成", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    {
                        Process.Start("explorer.exe", $"/select,\"{zipFullPath}\"");
                    }
                }
                catch (Exception ex) { MessageBox.Show("打包失败: " + ex.Message); }
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
            // GetAllSelectedItems 现在会收集 ProjectFileListItems 中 IsChecked 为 true 的项
            var items = GetAllSelectedItems();
            if (items.Count == 0)
            {
                var current = GetSelectedItem();
                if (current != null) items.Add(current);
            }

            foreach (var item in items)
            {
                if (item.Type == ExplorerItemType.File && item.FullPath.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
                {
                    CadService.InsertDwgAsBlock(item.FullPath); //
                }
            }
        }

        // 文件位置：sir-sz/cadatlasmanager/CadAtlasManager-FZ1/CadAtlasManager/UI/AtlasView.xaml.cs

        private void LoadConfig()
        {
            try
            {
                var config = ConfigManager.Load();
                if (config == null) return;

                // --- 原有逻辑：加载图集文件夹和项目列表 ---
                if (config.AtlasFolders != null)
                    foreach (var f in config.AtlasFolders)
                        if (Directory.Exists(f) && !_loadedAtlasFolders.Contains(f))
                        { _loadedAtlasFolders.Add(f); AddFolderToTree(f); }

                if (config.Projects != null)
                    foreach (var p in config.Projects)
                        if (Directory.Exists(p.Path)) ProjectList.Add(p);

                if (!string.IsNullOrEmpty(config.LastActiveProjectPath))
                {
                    var t = ProjectList.FirstOrDefault(p => p.Path == config.LastActiveProjectPath);
                    if (t != null) CbProjects.SelectedItem = t;
                }

                // --- 【新增代码】恢复布局宽度 ---
                ApplySavedLayout(config);
            }
            catch { }
        }

        // --- 【新增方法】将保存的数值应用到界面 ---
        private void ApplySavedLayout(AppConfig config)
        {
            // 1. 恢复项目工作台布局
            // 恢复左侧树宽度 (GridColumn 需要使用 GridLength)
            if (config.ProjectTreeWidth > 50)
                ColProjectTree.Width = new GridLength(config.ProjectTreeWidth);

            // 恢复右侧列表“文件名称”列宽
            //if (config.ProjectNameColumnWidth > 50)
            // ColProjectFileName.Width = config.ProjectNameColumnWidth;

            // 2. 恢复图纸工作台布局
            if (config.PlotTreeWidth > 50)
                ColPlotTree.Width = new GridLength(config.PlotTreeWidth);

            //if (config.PlotNameColumnWidth > 50)
            //ColPlotFileName.Width = config.PlotNameColumnWidth;
        }

        private void SaveConfig() { string ap = _activeProject?.Path ?? ""; ConfigManager.Save(_loadedAtlasFolders, ProjectList, ap); }

        private void BtnLoadFolder_Click(object sender, RoutedEventArgs e) { using (var d = new WinForms.FolderBrowserDialog()) { if (d.ShowDialog() == WinForms.DialogResult.OK && !_loadedAtlasFolders.Contains(d.SelectedPath)) { _loadedAtlasFolders.Add(d.SelectedPath); AddFolderToTree(d.SelectedPath); SaveConfig(); } } }
        private void BtnAddProject_Click(object sender, RoutedEventArgs e) { using (var d = new WinForms.FolderBrowserDialog()) { if (d.ShowDialog() == WinForms.DialogResult.OK && !ProjectList.Any(p => p.Path == d.SelectedPath)) { var p = new ProjectItem(Path.GetFileName(d.SelectedPath), d.SelectedPath); ProjectList.Add(p); CbProjects.SelectedItem = p; SaveConfig(); } } }

        private void ClearAllSelection(ObservableCollection<FileSystemItem> items) { if (items == null) return; foreach (var i in items) { i.IsItemSelected = false; ClearAllSelection(i.Children); } }
        // [修改文件: UI/AtlasView.xaml.cs]
        private List<FileSystemItem> GetAllSelectedItems()
        {
            var list = new List<FileSystemItem>();

            // 获取当前选中的选项卡索引：0-资料库, 1-项目, 2-图纸
            int tabIndex = MainTabControl.SelectedIndex;

            if (tabIndex == 0)
            {
                // 资料库保持原样
                CollectSelected(Items, list);
            }
            else if (tabIndex == 1)
            {
                // --- 核心修复：项目工作台 ---
                // 优先检查右侧列表（明细表）是否有勾选或选中的项
                var listSelected = ProjectFileListItems.Where(i => i.IsChecked || i.IsItemSelected).ToList();

                if (listSelected.Count > 0)
                {
                    // 如果右侧列表有选中内容，则只处理列表内容，不碰左侧树
                    list.AddRange(listSelected);
                }
                else
                {
                    // 只有当右侧列表完全没选时，才去收集左侧树的选中项（用于删除整个目录）
                    CollectSelected(ProjectTreeItems, list);
                }
            }
            else if (tabIndex == 2)
            {
                // --- 核心修复：图纸工作台 ---
                // 同样优先检查右侧列表
                var listSelected = PlotFileListItems.Where(i => i.IsChecked || i.IsItemSelected).ToList();

                if (listSelected.Count > 0)
                {
                    list.AddRange(listSelected);
                }
                else
                {
                    CollectSelected(PlotFolderItems, list);
                }
            }

            return list;
        }
        // --- 1. 图纸工作台右键触发：启动绑定流程 ---
        private void MenuItem_BindExternalPdf_Click(object sender, RoutedEventArgs e)
        {
            // 获取当前在图纸工作台选中的 PDF 项
            var item = PlotFileList.SelectedItem as FileSystemItem;
            if (item == null || !item.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("请先在列表中选择一个需要绑定的 PDF 文件。", "提示");
                return;
            }

            // 记录待绑定对象并切换到“项目工作台”
            _pendingBindingPdf = item;
            MainTabControl.SelectedIndex = 1;

            // 获取项目工作台的右键菜单并显示“确认绑定”相关的 UI
            var contextMenu = this.Resources["ProjectFileContextMenu"] as ContextMenu;
            if (contextMenu != null)
            {
                SetBindMenuVisibility(contextMenu, Visibility.Visible);
            }

            MessageBox.Show($"已进入绑定模式。\n\n当前待绑定文件: {item.Name}\n请在【项目工作台】找到对应的 DWG 源码文件，右键点击并选择【\u2705 确认绑定 PDF】。", "绑定引导");
        }

        // --- 2. 辅助方法：统一控制绑定菜单项及其分隔线的显示/隐藏 ---
        private void SetBindMenuVisibility(ContextMenu menu, Visibility visibility)
        {
            // 根据 Header 文字寻找“确认绑定”菜单项
            var bindMenu = menu.Items.OfType<MenuItem>().FirstOrDefault(m => m.Header != null && m.Header.ToString().Contains("确认绑定"));
            if (bindMenu != null) bindMenu.Visibility = visibility;

            // 寻找在 XAML 中定义的名为 BindSeparator 的分隔符
            var separator = menu.Items.OfType<Separator>().FirstOrDefault(s => s.Name == "BindSeparator");
            if (separator != null) separator.Visibility = visibility;
        }

        // --- 3. 项目工作台右键触发：执行最终绑定 ---
        private void MenuItem_ConfirmBind_Click(object sender, RoutedEventArgs e)
        {
            if (_pendingBindingPdf == null) return;

            // 获取项目工作台当前选中的 DWG 文件
            var dwgItem = ProjectFileList.SelectedItem as FileSystemItem;
            if (dwgItem == null || !dwgItem.Name.EndsWith(".dwg", StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("请选择一个 DWG 图纸文件进行绑定。", "提示");
                return;
            }

            string dwgPath = dwgItem.FullPath;
            string oldPdfPath = _pendingBindingPdf.FullPath;
            string pdfDir = Path.GetDirectoryName(oldPdfPath);

            // 逻辑：处理文件名一致性（一对一原则）
            string dwgNameNoExt = Path.GetFileNameWithoutExtension(dwgPath);
            string newPdfName = dwgNameNoExt + ".pdf";
            string finalPdfPath = Path.Combine(pdfDir, newPdfName);

            // 如果文件名不匹配，提示重命名以符合系统规范
            if (!newPdfName.Equals(_pendingBindingPdf.Name, StringComparison.OrdinalIgnoreCase))
            {
                var res = MessageBox.Show($"检测到文件名不一致。是否将 PDF 重命名为:\n{newPdfName}？\n\n(建议重命名以确保系统能自动追踪版本状态)",
                                          "绑定建议", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

                if (res == MessageBoxResult.Cancel) return;
                if (res == MessageBoxResult.Yes)
                {
                    try
                    {
                        if (File.Exists(finalPdfPath))
                        {
                            MessageBox.Show("绑定失败：目标目录下已存在同名 PDF 文件。");
                            return;
                        }
                        File.Move(oldPdfPath, finalPdfPath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("重命名文件失败: " + ex.Message);
                        return;
                    }
                }
                else
                {
                    finalPdfPath = oldPdfPath; // 用户选择不重命名，使用原始路径
                }
            }

            // 执行指纹提取与持久化记录
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;

                // 调用 SavePlotRecord，它会自动通过 CadService 静默提取 DWG 指纹并保存元数据
                PlotMetaManager.SavePlotRecord(dwgPath, finalPdfPath, true);

                // --- 收尾工作 ---
                _pendingBindingPdf = null;
                var contextMenu = this.Resources["ProjectFileContextMenu"] as ContextMenu;
                if (contextMenu != null)
                {
                    SetBindMenuVisibility(contextMenu, Visibility.Collapsed);
                }

                MessageBox.Show("外部 PDF 绑定成功！该文件现在已纳入版本校验系统。");

                // 刷新并切回图纸工作台
                RefreshPlotTree();
                MainTabControl.SelectedIndex = 2;
            }
            catch (Exception ex)
            {
                MessageBox.Show("绑定过程中发生错误: " + ex.Message);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void CollectSelected(ObservableCollection<FileSystemItem> s, List<FileSystemItem> r) { if (s == null) return; foreach (var i in s) { if (i.IsItemSelected) r.Add(i); CollectSelected(i.Children, r); } }
        private void SelectRange(ObservableCollection<FileSystemItem> root, FileSystemItem s, FileSystemItem e) { var l = new List<FileSystemItem>(); FlattenTree(root, l); int i1 = l.IndexOf(s), i2 = l.IndexOf(e); if (i1 != -1 && i2 != -1) for (int i = Math.Min(i1, i2); i <= Math.Max(i1, i2); i++) l[i].IsItemSelected = true; }
        private void FlattenTree(ObservableCollection<FileSystemItem> n, List<FileSystemItem> r) { foreach (var node in n) { r.Add(node); if (node.IsExpanded) FlattenTree(node.Children, r); } }
        private TreeViewItem GetTreeViewItemUnderMouse(DependencyObject e) { while (e != null && !(e is TreeViewItem)) e = VisualTreeHelper.GetParent(e); return e as TreeViewItem; }

        private void MenuItem_Remove_Click(object sender, RoutedEventArgs e) { if (FileTree.SelectedItem is FileSystemItem i && i.IsRoot) { _loadedAtlasFolders.Remove(i.FullPath); Items.Remove(i); SaveConfig(); } }
        // [修改方法: BtnRemoveProject_Click]
        private void BtnRemoveProject_Click(object sender, RoutedEventArgs e)
        {
            if (CbProjects.SelectedItem is ProjectItem p)
            {
                ProjectList.Remove(p);
                ProjectTreeItems.Clear();
                PlotFolderItems.Clear();
                PlotFileListItems.Clear();

                _activeProject = null;
                _currentProjectFolderPath = "";
                _currentPlotFolderPath = ""; // 【新增】同步重置

                SaveConfig();
            }
        }

        private void CbProjects_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CbProjects.SelectedItem is ProjectItem p)
            {
                _activeProject = p;

                // 1. 重置路径记录（确保“浏览文件”回到根目录）
                _currentProjectFolderPath = "";
                _currentPlotFolderPath = "";

                // 2. 【核心修改】显式清空右侧明细列表
                // 这样在点击新项目的文件夹之前，右侧面板会保持干净
                ProjectFileListItems.Clear();
                PlotFileListItems.Clear();

                // 3. 刷新左侧树结构
                RefreshProjectTree();
                RefreshPlotTree();

                SaveConfig();
            }
        }
        private void BtnRefreshProject_Click(object sender, RoutedEventArgs e) => RefreshProjectTree();
        // [修改方法: BtnRefresh_Click]
        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            // 根据当前选中的 Tab 索引执行刷新
            switch (MainTabControl.SelectedIndex)
            {
                case 0: // 资料库
                    ReloadAtlasTree();
                    break;
                case 1: // 项目工作台
                    RefreshProjectTree();
                    break;
                case 2: // 图纸工作台
                    RefreshPlotTree();
                    break;
            }
        }
        private void BtnRefreshPlot_Click(object sender, RoutedEventArgs e) => RefreshPlotTree();
        private void TbSearch_TextChanged(object sender, TextChangedEventArgs e) => ReloadAtlasTree();
        private void CbFileType_SelectionChanged(object sender, SelectionChangedEventArgs e) => ReloadAtlasTree();
        private void CbProjectFileType_SelectionChanged(object sender, SelectionChangedEventArgs e) => RefreshProjectTree();
        // [修改方法: BtnOpenProjectFolder_Click]
        private void BtnOpenProjectFolder_Click(object sender, RoutedEventArgs e)
        {
            if (_activeProject == null) return;

            // 默认打开项目根目录
            string targetPath = _activeProject.Path;

            // 逻辑：如果当前明细路径不为空，且目录确实存在，则打开明细目录
            // 这覆盖了“双击进入子目录”或“点击树节点”后的情况
            if (!string.IsNullOrEmpty(_currentProjectFolderPath) && Directory.Exists(_currentProjectFolderPath))
            {
                targetPath = _currentProjectFolderPath;
            }

            try
            {
                // 使用 Windows 资源管理器打开目标路径
                Process.Start("explorer.exe", targetPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开文件夹：\n{ex.Message}", "错误");
            }
        }
        // [修改方法: BtnOpenPlotFolder_Click]
        private void BtnOpenPlotFolder_Click(object sender, RoutedEventArgs e)
        {
            if (_activeProject == null) return;

            // 默认路径：项目定义的输出目录（通常是项目根目录下的 _Plot 文件夹）
            string targetPath = _activeProject.OutputPath;

            // 核心逻辑：如果当前图纸明细路径记录存在（即你点击或双击进入了某个子目录），则优先打开该目录
            if (!string.IsNullOrEmpty(_currentPlotFolderPath) && Directory.Exists(_currentPlotFolderPath))
            {
                targetPath = _currentPlotFolderPath;
            }

            try
            {
                // 调用资源管理器打开
                Process.Start("explorer.exe", targetPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开图纸目录：\n{ex.Message}", "错误");
            }
        }

        // [添加到 AtlasView.xaml.cs]
        // 获取当前所有已展开文件夹的路径
        private void GetExpandedPaths(ObservableCollection<FileSystemItem> nodes, List<string> expandedPaths)
        {
            foreach (var node in nodes)
            {
                if (node.Type == ExplorerItemType.Folder && node.IsExpanded)
                {
                    expandedPaths.Add(node.FullPath);
                    if (node.Children.Count > 0) GetExpandedPaths(node.Children, expandedPaths);
                }
            }
        }

        // 根据路径恢复展开和选中状态
        private void RestoreTreeState(ObservableCollection<FileSystemItem> nodes, List<string> expandedPaths, string selectedPath)
        {
            foreach (var node in nodes)
            {
                // 恢复展开状态
                if (expandedPaths.Contains(node.FullPath))
                {
                    node.IsExpanded = true;
                }

                // 恢复选中状态
                if (node.FullPath.Equals(selectedPath, StringComparison.OrdinalIgnoreCase))
                {
                    node.IsItemSelected = true;
                    // 确保右侧文件列表也加载该目录
                    LoadProjectFileListItems(node);
                }

                if (node.Children.Count > 0)
                {
                    RestoreTreeState(node.Children, expandedPaths, selectedPath);
                }
            }
        }
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

        // [修改方法：ValidatePdfVersion]
        private void ValidatePdfVersion(FileSystemItem item)
        {
            string plotDir = GetCurrentPlotDir();
            if (string.IsNullOrEmpty(plotDir)) return;

            // 获取基础状态
            PdfStatus status = GetPdfStatusRecursive(item.Name, plotDir);

            // 判定是否为外部绑定
            bool isExternal = PlotMetaManager.IsExternalBind(plotDir, item.Name);

            // 更新 UI
            UpdateVersionUi(item, status, isExternal);
        }

        // [修改方法：GetPdfStatusRecursive]
        // 2. 升级递归校验逻辑
        private PdfStatus GetPdfStatusRecursive(string pdfName, string plotDir)
        {
            if (PlotMetaManager.IsCombinedFile(plotDir, pdfName))
            {
                // 获取合并时记录的“指纹快照” (文件名:时间戳)
                var snapshots = PlotMetaManager.GetCombinedSourcesWithSnapshot(plotDir, pdfName);
                if (snapshots.Count == 0) return PdfStatus.Unknown;

                foreach (var entry in snapshots)
                {
                    string sourcePdfName = entry.Key;
                    string savedTimestamp = entry.Value;
                    string sourcePdfPath = Path.Combine(plotDir, sourcePdfName);

                    // A. 基础检查：文件是否还在磁盘上
                    if (!File.Exists(sourcePdfPath)) return PdfStatus.MissingSource;

                    // B. 时间认证（你的想法核心）：磁盘上的文件时间必须与合并时记录的时间一致
                    string currentTimestamp = File.GetLastWriteTimeUtc(sourcePdfPath).Ticks.ToString();
                    if (currentTimestamp != savedTimestamp)
                    {
                        // 虽然分项PDF可能是最新的，但它相对于合并文件来说是“新原材料”，需要重并
                        return PdfStatus.NeedRemerge;
                    }

                    // C. 状态认证：分项 PDF 自身是否过期（递归检查它对应的 DWG）
                    var subStatus = GetPdfStatusRecursive(sourcePdfName, plotDir);
                    if (subStatus != PdfStatus.Latest) return subStatus;
                }
                return PdfStatus.Latest;
            }
            else
            {
                return CheckSingleDwgStatus(pdfName, plotDir);
            }
        }

        // [修改方法：CheckSingleDwgStatus]
        // [修改文件: UI/AtlasView.xaml.cs]
        private PdfStatus CheckSingleDwgStatus(string pdfName, string plotDir)
        {
            // 1. 获取 DWG 文件名 (从记录或猜测)
            PlotMetaManager.LoadHistory(plotDir);
            string realDwgName = PlotMetaManager.GetSourceDwgName(plotDir, pdfName);

            if (string.IsNullOrEmpty(realDwgName))
            {
                string baseName = System.Text.RegularExpressions.Regex.Replace(Path.GetFileNameWithoutExtension(pdfName), @"_\d+$", "");
                realDwgName = baseName + ".dwg";
            }

            // 2. 【核心修改】定位 DWG 路径：仅搜索 plotDir (_Plot) 的上一级目录
            // 假设结构为：Folder/Sub/Drawing.dwg 和 Folder/Sub/_Plot/Drawing.pdf
            string sourceFolder = Path.GetDirectoryName(plotDir); // 获取 _Plot 的父目录
            string sourceDwgPath = Path.Combine(sourceFolder, realDwgName);

            // 检查文件是否存在于该特定目录下，不再进行全项目递归
            if (!File.Exists(sourceDwgPath)) return PdfStatus.MissingSource;

            // 3. 校验指纹
            string currentTimestamp = CadService.GetFileTimestamp(sourceDwgPath);
            bool isLatest = PlotMetaManager.CheckStatus(plotDir, pdfName, Path.GetFileName(sourceDwgPath), currentTimestamp,
                () => CadService.GetContentFingerprint(sourceDwgPath));

            return isLatest ? PdfStatus.Latest : PdfStatus.Expired;
        }

        // 递归核心逻辑
        private PdfStatus GetPdfStatusRecursive(string pdfName)
        {
            string plotDir = _activeProject.OutputPath;

            // 1. 判断是否为合并文件
            if (PlotMetaManager.IsCombinedFile(plotDir, pdfName))
            {
                // --- 合并文件逻辑 ---
                var sources = PlotMetaManager.GetCombinedSources(plotDir, pdfName);
                if (sources == null || sources.Count == 0) return PdfStatus.Unknown;

                bool anyExpired = false;
                bool anyMissing = false;

                foreach (var sourcePdf in sources)
                {
                    // 递归检查每一个源 PDF 的状态
                    var subStatus = GetPdfStatusRecursive(sourcePdf);

                    if (subStatus == PdfStatus.MissingSource) anyMissing = true;
                    if (subStatus == PdfStatus.Expired) anyExpired = true;
                }

                // 判定规则：
                // 1. 如果有任何源文件找不到 -> 源缺失
                // 2. 如果有任何源文件过期 -> 整体过期 (一票否决)
                // 3. 否则 -> 最新
                if (anyMissing) return PdfStatus.MissingSource;
                if (anyExpired) return PdfStatus.Expired;
                return PdfStatus.Latest;
            }
            else
            {
                // --- 普通文件逻辑 (DWG -> PDF) ---
                return CheckSingleDwgStatus(pdfName);
            }
        }

        // 单个 DWG 状态检查 (从原 ValidatePdfVersion 提取并优化)
        private PdfStatus CheckSingleDwgStatus(string pdfName)
        {
            // 1. 查名字
            string realDwgName = PlotMetaManager.GetSourceDwgName(_activeProject.OutputPath, pdfName);

            // 猜名字
            if (string.IsNullOrEmpty(realDwgName))
            {
                string baseName = System.Text.RegularExpressions.Regex.Replace(Path.GetFileNameWithoutExtension(pdfName), @"_\d+$", "");
                realDwgName = baseName + ".dwg";
            }

            // 2. 找文件
            string sourceDwgPath = Path.Combine(_activeProject.Path, realDwgName);
            if (!File.Exists(sourceDwgPath))
            {
                // 递归搜
                try
                {
                    string onlyFileName = Path.GetFileName(realDwgName);
                    var foundFiles = Directory.GetFiles(_activeProject.Path, onlyFileName, SearchOption.AllDirectories);
                    if (foundFiles.Length > 0) sourceDwgPath = foundFiles[0];
                    else return PdfStatus.MissingSource; // 彻底找不到
                }
                catch { return PdfStatus.MissingSource; }
            }

            // 3. 校验指纹
            string currentTimestamp = CadService.GetFileTimestamp(sourceDwgPath);
            bool isLatest = PlotMetaManager.CheckStatus(_activeProject.OutputPath, pdfName, Path.GetFileName(sourceDwgPath), currentTimestamp,
                () => CadService.GetContentFingerprint(sourceDwgPath));

            return isLatest ? PdfStatus.Latest : PdfStatus.Expired;
        }
        // [添加到 AtlasView.xaml.cs]
        // [AtlasView.xaml.cs]
        private string GetCurrentPlotDir()
        {
            // 1. 核心修复：优先使用成员变量记录的路径，这在 LoadPlotFilesList 一开始就赋值了
            // 这样能确保即使列表还是空的，也能正确识别 _Plot 目录
            string path = _currentPlotFolderPath;

            if (string.IsNullOrEmpty(path))
            {
                // 兜底逻辑：如果记录为空，再尝试从列表或树节点获取
                var item = PlotFileListItems.FirstOrDefault();
                if (item == null)
                {
                    item = PlotFolderTree.SelectedItem as FileSystemItem;
                }
                if (item != null) path = item.FullPath;
            }

            if (string.IsNullOrEmpty(path)) return null;

            // 向上寻找路径中包含 _Plot 的部分
            int idx = path.IndexOf("_Plot", StringComparison.OrdinalIgnoreCase);
            if (idx != -1)
            {
                return path.Substring(0, idx + 5);
            }
            return null;
        }

        // =================================================================
        // 【核心修复】右键菜单 - 打开源 DWG (同步最新的查找逻辑)
        // =================================================================
        // [修改文件: UI/AtlasView.xaml.cs]
        private void MenuItem_OpenSourceDwg_Click(object sender, RoutedEventArgs e)
        {
            var item = PlotFileList.SelectedItem as FileSystemItem;
            if (item == null || !item.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) return;

            string plotDir = GetCurrentPlotDir();
            if (string.IsNullOrEmpty(plotDir)) return;

            // 确定 DWG 文件名
            string realDwgName = PlotMetaManager.GetSourceDwgName(plotDir, item.Name);
            if (string.IsNullOrEmpty(realDwgName))
            {
                string baseName = System.Text.RegularExpressions.Regex.Replace(Path.GetFileNameWithoutExtension(item.Name), @"_\d+$", "");
                realDwgName = baseName + ".dwg";
            }

            // 【核心修改】直接在 _Plot 的上一级目录查找
            string sourceDwgPath = Path.Combine(Path.GetDirectoryName(plotDir), realDwgName);

            if (File.Exists(sourceDwgPath))
            {
                CadService.OpenDwg(sourceDwgPath, "Edit");
            }
            else
            {
                MessageBox.Show($"在该 PDF 的相对路径（上一级目录）未找到源文件：\n{realDwgName}", "文件未找到");
            }
        }
        private void BtnVersion_Click(object sender, RoutedEventArgs e) => MessageBox.Show(_versionInfo);
        private void ProjectTree_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e) { _dragStartPoint = e.GetPosition(null); var t = GetTreeViewItemUnderMouse(e.OriginalSource as DependencyObject); if (t != null) _draggedItem = t.DataContext as FileSystemItem; OnTreeItemClick(sender, e); }
        private void ProjectTree_MouseMove(object sender, MouseEventArgs e) { if (e.LeftButton == MouseButtonState.Pressed && _draggedItem != null) { var diff = _dragStartPoint - e.GetPosition(null); if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance || Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance) { var s = GetAllSelectedItems(); if (!s.Contains(_draggedItem)) { s.Clear(); s.Add(_draggedItem); } DataObject d = new DataObject(DataFormats.FileDrop, s.Select(i => i.FullPath).ToArray()); DragDrop.DoDragDrop(ProjectTree, d, DragDropEffects.Copy | DragDropEffects.Move); _draggedItem = null; } } }
        private void MenuItem_Open_Click(object sender, RoutedEventArgs e) { var i = GetSelectedItem(); if (i != null && i.Type == ExplorerItemType.File) ExecuteFileAction(i); }
        // 更新获取选中项的逻辑，确保 DataGrid 里的文件优先级最高
        private FileSystemItem GetSelectedItem()
        {
            int tabIndex = MainTabControl.SelectedIndex;

            switch (tabIndex)
            {
                case 0: // 资料图集库
                    return FileTree.SelectedItem as FileSystemItem;
                case 1: // 项目工作台
                    if (ProjectFileList.SelectedItem != null) return ProjectFileList.SelectedItem as FileSystemItem;
                    return ProjectTree.SelectedItem as FileSystemItem;
                case 2: // 图纸工作台
                    if (PlotFileList.SelectedItem != null) return PlotFileList.SelectedItem as FileSystemItem;
                    return PlotFolderTree.SelectedItem as FileSystemItem;
                default:
                    return _lastSelectedItem;
            }
        }
        private void MenuItem_NewFolder_Click(object sender, RoutedEventArgs e) { var i = GetSelectedItem(); if (i != null && i.Type == ExplorerItemType.Folder) { Directory.CreateDirectory(Path.Combine(i.FullPath, "新建文件夹")); RefreshProjectTree(); } }
        private void MenuItem_Expand_Click(object sender, RoutedEventArgs e)
        {
            // 强制获取 TreeView 的选中项，而不是 DataGrid 的
            FileSystemItem item = ProjectTree.SelectedItem as FileSystemItem ?? GetSelectedItem();
            if (item != null) SetExpansion(item, true);
        }
        private void MenuItem_Collapse_Click(object sender, RoutedEventArgs e)
        {
            FileSystemItem item = ProjectTree.SelectedItem as FileSystemItem ?? GetSelectedItem();
            if (item != null) SetExpansion(item, false);
        }
        // 确保递归逻辑能穿透到所有文件夹层级
        private void SetExpansion(FileSystemItem i, bool e)
        {
            i.IsExpanded = e;
            // 递归处理所有子项
            foreach (var c in i.Children)
            {
                if (c.Type == ExplorerItemType.Folder)
                    SetExpansion(c, e);
            }
        }
        private void FileTree_MouseDoubleClick(object sender, MouseButtonEventArgs e) { if (GetClickedItem(e) is FileSystemItem i && i.Type == ExplorerItemType.File) ExecuteFileAction(i); }
        // 项目工作台的目录树双击：也强制 Edit
        private void ProjectTree_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (GetClickedItem(e) is FileSystemItem item && item.Type == ExplorerItemType.File)
            {
                OpenFileSmart(item.FullPath, "Edit");
            }
        }
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
                // 2. 【核心修复】清除“项目工作台”右侧列表的勾选
                if (ProjectFileListItems != null)
                {
                    foreach (var item in ProjectFileListItems) item.IsChecked = false;
                }

                // 3. 【保留】清除“图纸工作台”右侧列表的勾选
                if (PlotFileListItems != null)
                {
                    foreach (var item in PlotFileListItems) item.IsChecked = false;
                }

                // 4. 执行当前项的单选操作
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
            if (f == "图片" && ".jpg.jpeg.png.bmp.gif.tif.tiff".Contains(x)) return true;
            if (f == "PDF" && x == ".pdf") return true;
            if (f == "压缩包" && ".zip.rar.7z".Contains(x)) return true;
            return false;
        }
        private string GetIconForExtension(string x)
        {
            if (x.Contains("dwg")) return "\uD83D\uDCD0";
            if (x == ".pat") return "\uD83E\uDD93"; // 斑马纹图标，很形象地代表填充图案
            if (".doc.docx.xls.xlsx.ppt.pptx.wps.txt".Contains(x)) return "\uD83D\uDCC4";
            if (".jpg.jpeg.png.bmp.gif.tif.tiff".Contains(x)) return "\uD83D\uDDBC\uFE0F";
            if (x.Contains("pdf")) return "\uD83D\uDCD5";
            if (".zip.rar.7z".Contains(x)) return "\uD83D\uDCE6";
            return "\uD83D\uDCC3";
        }
        private void BtnSaveLayout_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var config = ConfigManager.Load();

                // 记录项目工作台宽度
                config.ProjectTreeWidth = ColProjectTree.ActualWidth;
                config.ProjectNameColumnWidth = ColProjectFileName.ActualWidth;

                // 记录图纸工作台宽度
                config.PlotTreeWidth = ColPlotTree.ActualWidth;
                config.PlotNameColumnWidth = ColPlotFileName.ActualWidth;

                if (MainPlugin._ps != null)
                {
                    // 记录面板当前的宽度和高度
                    config.PaletteWidth = MainPlugin._ps.Size.Width;
                    config.PaletteHeight = MainPlugin._ps.Size.Height;
                }

                ConfigManager.Save(config);
                MessageBox.Show("布局与窗口尺寸已保存！", "提示");
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存失败: " + ex.Message);
            }
        }

    }
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