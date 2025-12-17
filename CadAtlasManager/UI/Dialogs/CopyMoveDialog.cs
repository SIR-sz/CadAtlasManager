using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace CadAtlasManager.UI
{
    public class CopyMoveDialog : Window
    {
        public string SelectedPath { get; private set; }
        public string FileName { get; private set; }

        private TextBox _tbPath;
        private TextBox _tbName;
        private bool _isMultiSelect;

        public CopyMoveDialog(string defaultPath, string defaultName, bool isMultiSelect)
        {
            _isMultiSelect = isMultiSelect;
            SelectedPath = defaultPath;
            FileName = defaultName;

            Width = 450; Height = 280;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = Brushes.WhiteSmoke;
            Title = "复制移动";

            var stackPanel = new StackPanel { Margin = new Thickness(20) };
            this.Content = stackPanel;

            // 1. 文件名区域
            stackPanel.Children.Add(new TextBlock
            {
                Text = isMultiSelect ? "文件名 (批量模式不可修改):" : "文件名 (重命名):",
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 5)
            });

            _tbName = new TextBox
            {
                Text = isMultiSelect ? "(保持原文件名)" : defaultName,
                Height = 28,
                VerticalContentAlignment = VerticalAlignment.Center,
                IsEnabled = !isMultiSelect,
                Foreground = isMultiSelect ? Brushes.Gray : Brushes.Black,
                Padding = new Thickness(2)
            };
            stackPanel.Children.Add(_tbName);

            stackPanel.Children.Add(new FrameworkElement { Height = 15 });

            // 2. 保存路径区域
            stackPanel.Children.Add(new TextBlock
            {
                Text = "保存路径:",
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 5)
            });

            var pathDock = new DockPanel { LastChildFill = true };

            var btnBrowse = new Button { Content = "...", Width = 35, Height = 28, Margin = new Thickness(5, 0, 0, 0) };
            DockPanel.SetDock(btnBrowse, Dock.Right);

            btnBrowse.Click += (s, e) =>
            {
                // 使用 fully qualified name 避免冲突
                using (var d = new System.Windows.Forms.FolderBrowserDialog())
                {
                    d.SelectedPath = SelectedPath;
                    if (d.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        _tbPath.Text = d.SelectedPath;
                    }
                }
            };

            _tbPath = new TextBox
            {
                Text = defaultPath,
                Height = 28,
                VerticalContentAlignment = VerticalAlignment.Center,
                Padding = new Thickness(2)
            };

            pathDock.Children.Add(btnBrowse);
            pathDock.Children.Add(_tbPath);
            stackPanel.Children.Add(pathDock);

            // 3. 底部按钮区域
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 30, 0, 0)
            };

            var btnOk = new Button { Content = "确定复制", Width = 90, Height = 32, IsDefault = true, Margin = new Thickness(0, 0, 10, 0) };
            btnOk.Click += (s, e) =>
            {
                SelectedPath = _tbPath.Text.Trim();
                if (!_isMultiSelect) FileName = _tbName.Text.Trim();

                if (string.IsNullOrEmpty(SelectedPath)) { MessageBox.Show("请选择保存路径"); return; }
                if (!_isMultiSelect && string.IsNullOrEmpty(FileName)) { MessageBox.Show("请输入文件名"); return; }

                DialogResult = true;
            };

            var btnCancel = new Button { Content = "取消", Width = 80, Height = 32, IsCancel = true };

            btnPanel.Children.Add(btnOk);
            btnPanel.Children.Add(btnCancel);
            stackPanel.Children.Add(btnPanel);
        }
    }
}