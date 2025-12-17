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

            // 1. 提示文本
            var txtInfo = new TextBlock
            {
                Text = "检测到以下文件的版本状态。请勾选需要打印的文件：\n(注：'已修改' 建议重新打印，'未修改' 可跳过)",
                Foreground = Brushes.Gray,
                Margin = new Thickness(0, 0, 0, 10)
            };
            grid.Children.Add(txtInfo);

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
                DialogResult = true;
            };

            var btnCancel = new Button { Content = "取消", Width = 80, Height = 30, IsCancel = true };

            btnPanel.Children.Add(btnCancel);
            btnPanel.Children.Add(btnPrint);
            grid.Children.Add(btnPanel);
        }
    }
}