using Autodesk.AutoCAD.ApplicationServices;
using CadAtlasManager.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace CadAtlasManager.UI
{
    public partial class BatchPlotDialog : Window
    {
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
                OffsetY = offY
            };
        }

        // --- 核心修改2：开始批量自动打印 ---
        private async void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            // --- 新增：确保打印前有活动窗口 ---
            CadService.EnsureHasActiveDocument();
            // 启动前再次确保环境干净
            CadService.ForceCleanup();

            var config = GetCurrentUiConfig();
            if (config == null) return;

            var selectedItems = _candidates.Where(c => c.IsSelected).ToList();
            if (selectedItems.Count == 0) return;

            var btn = sender as Button;
            btn.IsEnabled = false; // 禁用按钮，防止打印中途再次点击

            foreach (var item in selectedItems)
            {
                // 1. 更新状态文字，此时还在主线程
                item.Status = "正在打印...";


                // --- 核心修复：自动滚动到当前正在打印的项目 ---
                LvFiles.ScrollIntoView(item);

                // 2. 重要：给 UI 线程刷新时间
                await Task.Delay(100);

                try
                {
                    string directory = System.IO.Path.GetDirectoryName(item.FilePath);
                    string outputDir = System.IO.Path.Combine(directory, "_Plot");

                    if (!System.IO.Directory.Exists(outputDir))
                        System.IO.Directory.CreateDirectory(outputDir);

                    // 【核心修复】：移除 Task.Run，直接在主线程调用
                    // 虽然打印单个文件时 UI 会暂时无响应，但这在 AutoCAD 中是最安全、不会崩溃的做法
                    var results = CadService.BatchPlotByTitleBlocks(item.FilePath, outputDir, config);

                    if (results != null && results.Count > 0)
                    {
                        item.IsSuccess = true;
                        item.Status = $"成功({results.Count}张)";
                    }
                    else
                    {
                        item.IsSuccess = false;
                        item.Status = "未识别到图框";
                    }
                }
                catch (Exception ex)
                {
                    item.IsSuccess = false;
                    item.Status = "错误";
                    // 在主线程中，可以安全地弹出报错信息
                    MessageBox.Show($"文件 {item.FileName} 打印失败：\n{ex.Message}");
                }

                // 3. 处理完一个文件后，再次短暂释放 CPU，让界面更新
                await Task.Delay(50);
            }

            btn.IsEnabled = true;
            SaveCurrentConfig(config);
        }

        // --- 核心修改3：手动打印按钮点击事件 ---
        private void BtnManualPlot_Click(object sender, RoutedEventArgs e)
        {
            // --- 新增：确保开始前有背景窗口 ---
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

                // --- 核心修复：在这里定义 outputDir ---
                string directory = System.IO.Path.GetDirectoryName(item.FilePath);
                string outputDir = System.IO.Path.Combine(directory, "_Plot");

                if (!System.IO.Directory.Exists(outputDir))
                    System.IO.Directory.CreateDirectory(outputDir);

                // 调用手动拾取服务，务必传递 outputDir
                int count = CadService.ManualPickAndPlot(doc, outputDir, config);

                if (count > 0)
                {
                    item.IsSuccess = true;
                    item.Status = $"手动成功({count}张)";
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