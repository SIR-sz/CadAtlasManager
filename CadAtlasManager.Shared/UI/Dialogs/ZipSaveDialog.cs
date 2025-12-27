using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace CadAtlasManager.UI
{
    public class ZipSaveDialog : Window
    {
        public string SelectedPath { get; private set; }
        public string FileName { get; private set; }

        private TextBox _tbPath;
        private TextBox _tbName;

        public ZipSaveDialog(string defaultPath, string defaultName)
        {
            SelectedPath = defaultPath;
            FileName = defaultName;

            Width = 450; Height = 280;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            Background = Brushes.WhiteSmoke;
            Title = "一键打包";

            var stackPanel = new StackPanel { Margin = new Thickness(20) };
            this.Content = stackPanel;

            // 1. 文件名区域
            stackPanel.Children.Add(new TextBlock
            {
                Text = "压缩包名称:",
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 5)
            });

            _tbName = new TextBox
            {
                Text = defaultName,
                Height = 28,
                VerticalContentAlignment = VerticalAlignment.Center,
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
                // 使用 fully qualified name
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

            var btnOk = new Button { Content = "开始打包", Width = 90, Height = 32, IsDefault = true, Margin = new Thickness(0, 0, 10, 0) };
            btnOk.Click += (s, e) =>
            {
                SelectedPath = _tbPath.Text.Trim();
                FileName = _tbName.Text.Trim();

                if (string.IsNullOrEmpty(SelectedPath)) { MessageBox.Show("请选择保存路径"); return; }
                if (string.IsNullOrEmpty(FileName)) { MessageBox.Show("请输入压缩包名称"); return; }

                // 1. 拼接最终的压缩包完整路径
                string finalZipPath = System.IO.Path.Combine(SelectedPath, FileName + ".zip");

                try
                {
                    // 2. 调用压缩服务
                    // 这里的 sourceDir 是你要压缩的文件夹变量名
                    // CadAtlasManager.Services.ZipService.CreateZip(sourceDir, finalZipPath);

                    DialogResult = true;
                    MessageBox.Show("打包完成！");
                    this.Close();
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show("打包失败：" + ex.Message);
                }
            };

            var btnCancel = new Button { Content = "取消", Width = 80, Height = 32, IsCancel = true };

            btnPanel.Children.Add(btnOk);
            btnPanel.Children.Add(btnCancel);
            stackPanel.Children.Add(btnPanel);
        }
    }
}