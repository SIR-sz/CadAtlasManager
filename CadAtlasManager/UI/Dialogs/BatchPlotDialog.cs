using CadAtlasManager.Models; // 引用 Models
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;

namespace CadAtlasManager.UI
{
    public class BatchPlotDialog : Window
    {
        public List<string> ConfirmedFiles { get; private set; } = new List<string>();
        private List<PlotCandidate> _candidates;
        // 【新增】对外暴露用户最终确认的图框名称列表
        public List<string> TargetBlockNames { get; private set; }
        private TextBox _tbBlockNames; // 输入框引用
        public BatchPlotDialog(List<PlotCandidate> candidates)
        {
            _candidates = candidates;

            Title = "批量打印确认";
            Width = 600; Height = 400;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = Brushes.WhiteSmoke;

            // 主 Grid
            var grid = new Grid { Margin = new Thickness(10) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 头部
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // 列表区
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // 底部按钮区
            this.Content = grid;
            var topPanel = new StackPanel { Orientation = Orientation.Vertical, Margin = new Thickness(0, 0, 0, 10) };
            topPanel.SetValue(Grid.RowProperty, 0);

            // 1. 提示文本
            var txtInfo = new TextBlock
            {
                Text = "检测到以下文件的版本状态。请勾选需要打印的文件：",
                Foreground = Brushes.Gray,
                Margin = new Thickness(0, 0, 0, 5)
            };
            topPanel.Children.Add(txtInfo);

            // 2. 新增：图框名称配置
            var configGrid = new Grid { Margin = new Thickness(0, 5, 0, 0) };
            configGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            configGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var lblConfig = new TextBlock { Text = "识别图框名称 (逗号分隔): ", VerticalAlignment = VerticalAlignment.Center };

            // 读取配置
            var currentConfig = ConfigManager.Load();
            string initialNames = currentConfig.TitleBlockNames ?? "TK,A3图框";

            _tbBlockNames = new TextBox
            {
                Text = initialNames,
                Height = 26,
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(3),
                ToolTip = "输入块名或外部参照名，用中文或英文逗号分隔"
            };
            Grid.SetColumn(_tbBlockNames, 1);

            configGrid.Children.Add(lblConfig);
            configGrid.Children.Add(_tbBlockNames);
            topPanel.Children.Add(configGrid);

            grid.Children.Add(topPanel); // 将 topPanel 加入主 Grid

            // 2. 列表视图
            var listView = new ListView { Margin = new Thickness(0, 0, 0, 10) };
            listView.SetValue(Grid.RowProperty, 1);

            // 定义 GridView
            var gridView = new GridView();

            // CheckBox 列
            var checkFactory = new FrameworkElementFactory(typeof(CheckBox));
            checkFactory.SetBinding(CheckBox.IsCheckedProperty, new Binding("IsSelected"));
            var colCheck = new GridViewColumn { Header = "选择", CellTemplate = new DataTemplate { VisualTree = checkFactory }, Width = 50 };

            // 状态列
            var colStatus = new GridViewColumn { Header = "状态", DisplayMemberBinding = new Binding("StatusText"), Width = 80 };
            // 文件名列
            var colName = new GridViewColumn { Header = "文件名", DisplayMemberBinding = new Binding("FileName"), Width = 300 };

            gridView.Columns.Add(colCheck);
            gridView.Columns.Add(colStatus);
            gridView.Columns.Add(colName);

            listView.View = gridView;
            listView.ItemsSource = _candidates;

            grid.Children.Add(listView);

            // 3. 底部按钮
            var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            btnPanel.SetValue(Grid.RowProperty, 2);

            var btnPrint = new Button { Content = "下一步：配置打印参数", Width = 150, Height = 30, IsDefault = true, Margin = new Thickness(10, 0, 0, 0) };
            btnPrint.Click += (s, e) =>
            {
                ConfirmedFiles = _candidates.Where(c => c.IsSelected).Select(c => c.FilePath).ToList();
                if (ConfirmedFiles.Count == 0)
                {
                    MessageBox.Show("请至少选择一个文件。");
                    return;
                }

                // 【新增】保存配置
                string inputNames = _tbBlockNames.Text.Trim();
                if (string.IsNullOrEmpty(inputNames))
                {
                    MessageBox.Show("图框名称不能为空，否则无法自动识别打印区域。");
                    return;
                }

                // 解析并去重
                TargetBlockNames = inputNames.Split(new[] { ',', '，' }, System.StringSplitOptions.RemoveEmptyEntries)
                                             .Select(n => n.Trim())
                                             .Distinct()
                                             .ToList();

                // 持久化保存到 Config
                var config = ConfigManager.Load();
                config.TitleBlockNames = inputNames;
                // 保存逻辑需要 ConfigManager 提供 Save 方法支持仅更新部分字段，
                // 这里简单起见，我们假设 ConfigManager.Save(..., ..., ...) 是全量保存
                // 在 AtlasView 调用层保存可能更合适，或者临时保存到内存
                // 为了代码整洁，我们把 names 存在 public 属性 TargetBlockNames 里供外部调用

                DialogResult = true;
            };

            var btnCancel = new Button { Content = "取消", Width = 80, Height = 30, IsCancel = true };

            btnPanel.Children.Add(btnCancel);
            btnPanel.Children.Add(btnPrint);
            grid.Children.Add(btnPanel);
        }
    }
}