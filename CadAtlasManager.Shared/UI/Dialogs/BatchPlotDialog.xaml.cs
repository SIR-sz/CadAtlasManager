using Autodesk.AutoCAD.ApplicationServices;
using CadAtlasManager.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace CadAtlasManager.UI
{
    public partial class BatchPlotDialog : Window
    {
        // 1. 定义取消标记
        private bool _isCancelRequested = false;
        // --- 新增：标识是否为重印模式 ---
        public bool IsReprintMode { get; set; } = false;
        private List<PlotCandidate> _candidates;

        // 供外部获取最终结果（如果需要）
        public BatchPlotConfig FinalConfig { get; private set; }
        public List<string> ConfirmedFiles { get; private set; }

        public BatchPlotDialog(List<PlotCandidate> candidates)
        {
            InitializeComponent();
            // --- 核心修改：确保窗口在打印过程中置顶 ---
            this.Topmost = true;
            _candidates = candidates;
            LvFiles.ItemsSource = _candidates;
            TxtCount.Text = $"共 {_candidates.Count} 个文件";

            LoadInitialData();
            UpdateScaleUiState();
        }

        // 3. “取消打印”按钮逻辑
        private void BtnCancelPlot_Click(object sender, RoutedEventArgs e)
        {
            _isCancelRequested = true;
            BtnCancelPlot.IsEnabled = false; // 立即禁用防止重复点击
        }
        private void LoadInitialData()
        {
            // --- 新增：确保此时 CAD 中至少有一个活动文档 ---
            CadService.EnsureHasActiveDocument();
            var config = ConfigManager.Load() ?? new AppConfig();
            TbBlockNames.Text = config.TitleBlockNames ?? "TK,A3图框";

            var plotters = CadService.GetPlotters();
            CbPrinters.ItemsSource = plotters;
            string lastPrinter = !string.IsNullOrEmpty(config.LastPrinter) ? config.LastPrinter : "DWG To PDF.pc3";
            if (plotters.Contains(lastPrinter)) CbPrinters.SelectedItem = lastPrinter;
            else if (plotters.Count > 0) CbPrinters.SelectedIndex = 0;

            var styles = CadService.GetStyleSheets();
            CbStyles.ItemsSource = styles;
            string lastStyle = !string.IsNullOrEmpty(config.LastStyleSheet) ? config.LastStyleSheet : "monochrome.ctb";
            if (styles.Contains(lastStyle)) CbStyles.SelectedItem = lastStyle;

            if (config.LastOrderType == PlotOrderType.Vertical) RbOrderV.IsChecked = true;
            else RbOrderH.IsChecked = true;
            // ============== 新增：填充比例下拉框预设值 ==============
            var scalePresets = new List<string>
            {
                "1:1",
                "1:2", "1:4", "1:5", "1:8", "1:10", "1:16", "1:20", "1:30", "1:40", "1:50", "1:100",
                "2:1", "4:1", "8:1", "10:1", "100:1"
            };
            CbScale.ItemsSource = scalePresets;

            ChkFitToPaper.IsChecked = config.LastFitToPaper;
            CbScale.Text = !string.IsNullOrEmpty(config.LastScaleType) ? config.LastScaleType : "1:1";
            ChkCenterPlot.IsChecked = config.LastCenterPlot;
            TbOffsetX.Text = config.LastOffsetX.ToString();
            TbOffsetY.Text = config.LastOffsetY.ToString();
            ChkAutoRotate.IsChecked = config.LastAutoRotate;

        }

        // --- 核心修改1：实时从 UI 抓取配置的方法 ---
        private BatchPlotConfig GetCurrentUiConfig()
        {
            if (CbPrinters.SelectedItem == null || CbPaper.SelectedItem == null) return null;

            double offX = 0, offY = 0;
            double.TryParse(TbOffsetX.Text, out offX);
            double.TryParse(TbOffsetY.Text, out offY);

            string scaleStr = ChkFitToPaper.IsChecked == true ? "Fit" : CbScale.Text.Trim();
            if (string.IsNullOrEmpty(scaleStr)) scaleStr = "1:1";

            return new BatchPlotConfig
            {
                PrinterName = CbPrinters.SelectedItem.ToString(),
                MediaName = CbPaper.SelectedItem.ToString(),
                StyleSheet = CbStyles.SelectedItem?.ToString(),
                TitleBlockNames = TbBlockNames.Text.Trim(),
                AutoRotate = ChkAutoRotate.IsChecked == true,
                OrderType = (RbOrderV.IsChecked == true) ? PlotOrderType.Vertical : PlotOrderType.Horizontal,
                ScaleType = scaleStr,
                PlotCentered = ChkCenterPlot.IsChecked == true,
                OffsetX = offX,
                OffsetY = offY,
                ForceUseSelectedMedia = ChkLockPaper.IsChecked == true,

                // --- 新增：读取 UI 上的打印选项 ---
                PlotWithLineweights = ChkPlotLineweights.IsChecked == true,
                PlotTransparency = ChkPlotTransparency.IsChecked == true,
                PlotWithPlotStyles = ChkPlotStyles.IsChecked == true
            };
        }

        // [BatchPlotDialog.xaml.cs]

        /// <summary>
        /// 刷新指定目录下的目录文档
        /// </summary>
        private void UpdatePrintDirectoryLog(string targetDir)
        {
            try
            {
                if (string.IsNullOrEmpty(targetDir)) return;

                // 1. 过滤成功项 (保持原有逻辑)
                var successfulItemsInDir = _candidates
                    .Where(c => c.IsSuccess == true)
                    .Where(c =>
                    {
                        try
                        {
                            string dwgDir = System.IO.Path.GetDirectoryName(c.FilePath);
                            string outDir = System.IO.Path.Combine(dwgDir, "_Plot");
                            return string.Equals(System.IO.Path.GetFullPath(outDir),
                                                 System.IO.Path.GetFullPath(targetDir),
                                                 StringComparison.OrdinalIgnoreCase);
                        }
                        catch { return false; }
                    })
                    .ToList();

                if (successfulItemsInDir.Count == 0) return;

                // 2. 构造日志内容 (保持原有逻辑)
                var logLines = new List<string>();
                for (int i = 0; i < successfulItemsInDir.Count; i++)
                {
                    var item = successfulItemsInDir[i];
                    int pageCount = ExtractPageCountFromStatus(item.Status);
                    // --- 核心修改：去除文件名后缀 ---
                    string nameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(item.FileName);

                    // 格式：序号 文件名(无后缀) 张数
                    logLines.Add($"{i + 1} {nameWithoutExt} {pageCount}");
                }

                // --- 核心修改：根据模式决定文件名 ---
                string fileName = IsReprintMode ? "打印目录重印部分.txt" : "打印目录.txt";
                string logPath = System.IO.Path.Combine(targetDir, fileName);

                // 3. 写入文件
                System.IO.File.WriteAllLines(logPath, logLines, System.Text.Encoding.UTF8);
            }
            catch { }
        }


        /// <summary>
        /// 从状态字符串中提取数字
        /// </summary>
        private int ExtractPageCountFromStatus(string status)
        {
            if (string.IsNullOrEmpty(status)) return 0;
            var match = System.Text.RegularExpressions.Regex.Match(status, @"\((\d+)张\)");
            if (match.Success && int.TryParse(match.Groups[1].Value, out int count))
            {
                return count;
            }
            return 0;
        }
        // --- 核心修改2：开始批量自动打印 ---
        private async void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            CadService.EnsureHasActiveDocument();

            var config = GetCurrentUiConfig();
            if (config == null) return;

            var selectedItems = _candidates.Where(c => c.IsSelected).ToList();
            if (selectedItems.Count == 0) return;

            // --- UI 状态切换 ---
            _isCancelRequested = false;
            BtnPrint.IsEnabled = false;
            BtnCancelPlot.IsEnabled = true;

            var affectedDirs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < selectedItems.Count; i++)
            {
                var item = selectedItems[i];

                // --- 检查取消标记 ---
                if (_isCancelRequested)
                {
                    item.IsSuccess = false;
                    item.Status = "任务取消";
                    continue;
                }

                item.Status = "正在打印...";
                LvFiles.ScrollIntoView(item);

#if CAD2012
                // .NET 4.0 环境：Thread.Sleep 会阻塞当前线程。
                // 如果这是在后台任务（Task.Run）中运行，这是安全的。
                System.Threading.Thread.Sleep(100);
#else
                    // .NET 4.5+ 环境：使用异步等待，不阻塞线程。
                    await System.Threading.Tasks.Task.Delay(100);
#endif

                try
                {
                    string directory = System.IO.Path.GetDirectoryName(item.FilePath);
                    string outputDir = System.IO.Path.Combine(directory, "_Plot");

                    if (!System.IO.Directory.Exists(outputDir))
                        System.IO.Directory.CreateDirectory(outputDir);

                    // --- 核心修复：直接调用，不再使用 Task.Run ---
                    // 这样它就在 AutoCAD 的主线程中安全运行
                    var results = CadService.BatchPlotByTitleBlocks(item.FilePath, outputDir, config);

                    if (results != null && results.Count > 0)
                    {
                        item.IsSuccess = true;
                        item.Status = $"成功({results.Count}张)";
                        affectedDirs.Add(outputDir);
                    }
                    else
                    {
                        item.IsSuccess = false;
                        item.Status = "未识别到图框";
                    }
                }
                catch (Exception)
                {
                    item.IsSuccess = false;
                    item.Status = "错误";
                    // 如果仍然报错，取消下面一行的注释查看具体报错原因
                    // MessageBox.Show($"文件 {item.FileName} 打印失败：\n{ex.Message}");
                }

#if CAD2012
                // .NET 4.0 环境：Thread.Sleep 会阻塞当前线程。
                // 如果这是在后台任务（Task.Run）中运行，这是安全的。
                System.Threading.Thread.Sleep(100);
#else
                    // .NET 4.5+ 环境：使用异步等待，不阻塞线程。
                    await System.Threading.Tasks.Task.Delay(100);
#endif
            }

            foreach (var dir in affectedDirs)
            {
                UpdatePrintDirectoryLog(dir);
            }

            BtnPrint.IsEnabled = true;
            BtnCancelPlot.IsEnabled = false;
            SaveCurrentConfig(config);
        }

        // --- 核心修改3：手动打印按钮点击事件 ---
        // [BatchPlotDialog.xaml.cs]
        private void BtnManualPlot_Click(object sender, RoutedEventArgs e)
        {
            // --- 确保开始前有背景窗口 ---
            CadService.EnsureHasActiveDocument();

            var btn = sender as Button;
            var item = btn.DataContext as PlotCandidate;
            if (item == null) return;

            var config = GetCurrentUiConfig();
            if (config == null) return;

            this.Hide(); // 隐藏界面以便操作 CAD

            try
            {
                var docMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
                var doc = docMgr.Open(item.FilePath, false);
                docMgr.MdiActiveDocument = doc;

                // 定义输出目录
                string directory = System.IO.Path.GetDirectoryName(item.FilePath);
                string outputDir = System.IO.Path.Combine(directory, "_Plot");

                if (!System.IO.Directory.Exists(outputDir))
                    System.IO.Directory.CreateDirectory(outputDir);

                // 调用手动拾取服务
                int count = CadService.ManualPickAndPlot(doc, outputDir, config);

                if (count > 0)
                {
                    // 1. 更新当前项的状态
                    item.IsSuccess = true;
                    item.Status = $"手动成功({count}张)";

                    // 2. 核心修改：立即刷新并重新排序生成该目录下的“打印目录.txt”
                    // 这样无论这个文件在列表中第几个，都会按列表顺序重新编号记录
                    UpdatePrintDirectoryLog(outputDir);
                }

                doc.CloseAndDiscard();
            }
            catch (Exception ex)
            {
                MessageBox.Show("手动打印出错: " + ex.Message);
            }
            finally
            {
                this.ShowDialog();
            }
        }

        private void SaveCurrentConfig(BatchPlotConfig final)
        {
            try
            {
                var config = ConfigManager.Load() ?? new AppConfig();
                config.TitleBlockNames = final.TitleBlockNames;
                config.LastPrinter = final.PrinterName;
                config.LastMedia = final.MediaName;
                config.LastStyleSheet = final.StyleSheet;
                config.LastOrderType = final.OrderType;
                config.LastAutoRotate = final.AutoRotate;
                config.LastFitToPaper = ChkFitToPaper.IsChecked == true;
                config.LastScaleType = CbScale.Text.Trim();
                config.LastCenterPlot = final.PlotCentered;
                config.LastOffsetX = final.OffsetX;
                config.LastOffsetY = final.OffsetY;

                // --- 如果你在 AppConfig 里加了字段，就在这里保存 ---
                // config.LastForceUseSelectedMedia = final.ForceUseSelectedMedia; 

                ConfigManager.Save(config);
            }
            catch { }
        }

        // --- 其他 UI 交互保持不变 ---
        private void CbPrinters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CbPrinters.SelectedItem == null) return;
            string device = CbPrinters.SelectedItem.ToString();
            var mediaList = CadService.GetMediaList(device);
            CbPaper.ItemsSource = mediaList;
            var config = ConfigManager.Load();
            string lastMedia = config?.LastMedia;
            if (!string.IsNullOrEmpty(lastMedia) && mediaList.Contains(lastMedia)) CbPaper.SelectedItem = lastMedia;
            else
            {
                var a3 = mediaList.FirstOrDefault(m => m.ToUpper().Contains("A3"));
                if (a3 != null) CbPaper.SelectedItem = a3;
                else if (mediaList.Count > 0) CbPaper.SelectedIndex = 0;
            }
        }

        private void ChkFitToPaper_CheckedChanged(object sender, RoutedEventArgs e) => UpdateScaleUiState();
        private void ChkCenterPlot_CheckedChanged(object sender, RoutedEventArgs e) => UpdateScaleUiState();

        private void UpdateScaleUiState()
        {
            if (CbScale == null || TbOffsetX == null) return;
            CbScale.IsEnabled = ChkFitToPaper.IsChecked != true;
            TbOffsetX.IsEnabled = TbOffsetY.IsEnabled = ChkCenterPlot.IsChecked != true;
        }

        private void HeaderCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox chk)
            {
                bool val = chk.IsChecked == true;
                foreach (var item in _candidates) item.IsSelected = val;
                LvFiles.Items.Refresh();
            }
        }

        // [BatchPlotDialog.xaml.cs] 修改 BtnClose_Click
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            // 1. 清空 UI 绑定数据源，断开引用链
            LvFiles.ItemsSource = null;
            _candidates = null;

            // 2. 执行核心清理逻辑：重置打印环境并强制 GC
            CadService.ForceCleanup();

            // 3. 关闭窗口
            this.DialogResult = true;
        }
    }
}