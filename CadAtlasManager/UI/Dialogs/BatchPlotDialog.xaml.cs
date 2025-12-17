using CadAtlasManager.Models;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace CadAtlasManager.UI
{
    public partial class BatchPlotDialog : Window
    {
        private List<PlotCandidate> _candidates;

        public BatchPlotConfig FinalConfig { get; private set; }
        public List<string> ConfirmedFiles { get; private set; }

        public BatchPlotDialog(List<PlotCandidate> candidates)
        {
            InitializeComponent();
            _candidates = candidates;
            LvFiles.ItemsSource = _candidates;
            TxtCount.Text = $"共 {_candidates.Count} 个文件";

            LoadInitialData();
            UpdateScaleUiState(); // 初始化状态
        }

        private void LoadInitialData()
        {
            var config = ConfigManager.Load() ?? new AppConfig();
            TbBlockNames.Text = config.TitleBlockNames ?? "TK,A3图框";

            // 打印机
            var plotters = CadService.GetPlotters();
            CbPrinters.ItemsSource = plotters;
            string lastPrinter = !string.IsNullOrEmpty(config.LastPrinter) ? config.LastPrinter : "DWG To PDF.pc3";
            if (plotters.Contains(lastPrinter)) CbPrinters.SelectedItem = lastPrinter;
            else if (plotters.Count > 0) CbPrinters.SelectedIndex = 0;

            // 样式表
            var styles = CadService.GetStyleSheets();
            CbStyles.ItemsSource = styles;
            string lastStyle = !string.IsNullOrEmpty(config.LastStyleSheet) ? config.LastStyleSheet : "monochrome.ctb";
            if (styles.Contains(lastStyle)) CbStyles.SelectedItem = lastStyle;
        }

        // 打印机改变，联动纸张
        private void CbPrinters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CbPrinters.SelectedItem == null) return;
            string device = CbPrinters.SelectedItem.ToString();
            var mediaList = CadService.GetMediaList(device);
            CbPaper.ItemsSource = mediaList;

            var config = ConfigManager.Load();
            string lastMedia = config?.LastMedia;
            if (!string.IsNullOrEmpty(lastMedia) && mediaList.Contains(lastMedia))
                CbPaper.SelectedItem = lastMedia;
            else
            {
                var a3 = mediaList.FirstOrDefault(m => m.ToUpper().Contains("A3"));
                if (a3 != null) CbPaper.SelectedItem = a3;
                else if (mediaList.Count > 0) CbPaper.SelectedIndex = 0;
            }
        }

        // 交互逻辑：勾选"布满"时，禁用比例输入框
        private void ChkFitToPaper_CheckedChanged(object sender, RoutedEventArgs e)
        {
            UpdateScaleUiState();
        }

        // 交互逻辑：勾选"居中"时，禁用偏移输入框
        private void ChkCenterPlot_CheckedChanged(object sender, RoutedEventArgs e)
        {
            UpdateScaleUiState();
        }

        private void UpdateScaleUiState()
        {
            if (CbScale == null || TbOffsetX == null) return;

            bool isFit = ChkFitToPaper.IsChecked == true;
            CbScale.IsEnabled = !isFit;

            bool isCenter = ChkCenterPlot.IsChecked == true;
            TbOffsetX.IsEnabled = !isCenter;
            TbOffsetY.IsEnabled = !isCenter;
        }

        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            ConfirmedFiles = _candidates.Where(c => c.IsSelected).Select(c => c.FilePath).ToList();
            if (ConfirmedFiles.Count == 0) { MessageBox.Show("请至少勾选一个文件。"); return; }
            if (CbPrinters.SelectedItem == null || CbPaper.SelectedItem == null) { MessageBox.Show("请选择打印机和纸张。"); return; }
            if (string.IsNullOrWhiteSpace(TbBlockNames.Text)) { MessageBox.Show("请输入图框块名。"); return; }

            // 解析偏移
            double offX = 0, offY = 0;
            double.TryParse(TbOffsetX.Text, out offX);
            double.TryParse(TbOffsetY.Text, out offY);

            // 解析比例
            // 如果勾选布满，则 Type="Fit"
            // 否则取下拉框的值（可能是 "1:100" 或 "1:1" 等）
            string scaleStr = "Fit";
            if (ChkFitToPaper.IsChecked == false)
            {
                scaleStr = CbScale.Text.Trim();
                if (string.IsNullOrEmpty(scaleStr)) scaleStr = "1:1"; // 默认兜底
            }

            FinalConfig = new BatchPlotConfig
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

            // 保存记忆
            try
            {
                var config = ConfigManager.Load() ?? new AppConfig();
                config.TitleBlockNames = FinalConfig.TitleBlockNames;
                config.LastPrinter = FinalConfig.PrinterName;
                config.LastMedia = FinalConfig.MediaName;
                config.LastStyleSheet = FinalConfig.StyleSheet;
                ConfigManager.Save(config);
            }
            catch { }

            DialogResult = true;
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
    }
}