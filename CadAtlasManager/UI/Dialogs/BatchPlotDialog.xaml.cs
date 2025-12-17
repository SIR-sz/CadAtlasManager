using CadAtlasManager.Models; // 引用模型
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace CadAtlasManager.UI
{
    public partial class BatchPlotDialog : Window
    {
        private List<PlotCandidate> _candidates;

        // 对外输出的配置结果
        public BatchPlotConfig FinalConfig { get; private set; }

        // 用户最终确认要打印的文件路径列表
        public List<string> ConfirmedFiles { get; private set; }

        public BatchPlotDialog(List<PlotCandidate> candidates)
        {
            InitializeComponent();
            _candidates = candidates;
            LvFiles.ItemsSource = _candidates;
            TxtCount.Text = $"共 {_candidates.Count} 个文件";

            // 初始化加载数据
            LoadInitialData();
        }

        private void LoadInitialData()
        {
            // 1. 读取历史配置
            var config = ConfigManager.Load() ?? new AppConfig();

            // 填入图框名
            TbBlockNames.Text = config.TitleBlockNames ?? "TK,A3图框";

            // 2. 获取 CAD 环境数据 (打印机列表)
            var plotters = CadService.GetPlotters();
            CbPrinters.ItemsSource = plotters;

            // 尝试选中上次使用的打印机，如果没有则选 PDF
            string lastPrinter = !string.IsNullOrEmpty(config.LastPrinter) ? config.LastPrinter : "DWG To PDF.pc3";

            if (plotters.Contains(lastPrinter))
            {
                CbPrinters.SelectedItem = lastPrinter;
            }
            else if (plotters.Count > 0)
            {
                CbPrinters.SelectedIndex = 0;
            }

            // 3. 获取样式表列表
            var styles = CadService.GetStyleSheets();
            CbStyles.ItemsSource = styles;

            string lastStyle = !string.IsNullOrEmpty(config.LastStyleSheet) ? config.LastStyleSheet : "monochrome.ctb";
            if (styles.Contains(lastStyle)) CbStyles.SelectedItem = lastStyle;
        }

        // 当打印机改变时，更新纸张列表
        private void CbPrinters_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CbPrinters.SelectedItem == null) return;
            string device = CbPrinters.SelectedItem.ToString();

            // 获取该设备支持的纸张
            var mediaList = CadService.GetMediaList(device);
            CbPaper.ItemsSource = mediaList;

            // 尝试选中上次使用的纸张，或者默认 A3
            var config = ConfigManager.Load();
            string lastMedia = config?.LastMedia;

            if (!string.IsNullOrEmpty(lastMedia) && mediaList.Contains(lastMedia))
            {
                CbPaper.SelectedItem = lastMedia;
            }
            else
            {
                // 智能默认：优先选 A3
                var a3 = mediaList.FirstOrDefault(m => m.ToUpper().Contains("A3"));
                if (a3 != null) CbPaper.SelectedItem = a3;
                else if (mediaList.Count > 0) CbPaper.SelectedIndex = 0;
            }
        }

        // 表头全选框逻辑
        private void HeaderCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox chk)
            {
                bool val = chk.IsChecked == true;
                foreach (var item in _candidates) item.IsSelected = val;
                // 刷新列表视图
                LvFiles.Items.Refresh();
            }
        }

        // 点击“开始打印”
        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            // 1. 基础校验
            ConfirmedFiles = _candidates.Where(c => c.IsSelected).Select(c => c.FilePath).ToList();
            if (ConfirmedFiles.Count == 0)
            {
                MessageBox.Show("请至少勾选一个要打印的文件。", "提示");
                return;
            }
            if (CbPrinters.SelectedItem == null || CbPaper.SelectedItem == null)
            {
                MessageBox.Show("请选择【打印机】和【默认纸张】。", "参数缺失");
                return;
            }
            if (string.IsNullOrWhiteSpace(TbBlockNames.Text))
            {
                MessageBox.Show("请输入【图框块名】，否则无法识别打印区域。", "参数缺失");
                return;
            }

            // 2. 生成配置对象
            FinalConfig = new BatchPlotConfig
            {
                PrinterName = CbPrinters.SelectedItem.ToString(),
                MediaName = CbPaper.SelectedItem.ToString(),
                StyleSheet = CbStyles.SelectedItem?.ToString(),
                TitleBlockNames = TbBlockNames.Text.Trim(),
                AutoRotate = ChkAutoRotate.IsChecked == true,
                UseStandardScale = true
            };

            // 3. 保存本次配置到磁盘 (记忆功能)
            try
            {
                var config = ConfigManager.Load() ?? new AppConfig();
                config.TitleBlockNames = FinalConfig.TitleBlockNames;
                config.LastPrinter = FinalConfig.PrinterName;
                config.LastMedia = FinalConfig.MediaName;
                config.LastStyleSheet = FinalConfig.StyleSheet;
                ConfigManager.Save(config);
            }
            catch { /* 保存失败不影响打印 */ }

            // 4. 关闭窗口，返回成功
            DialogResult = true;
        }
    }
}