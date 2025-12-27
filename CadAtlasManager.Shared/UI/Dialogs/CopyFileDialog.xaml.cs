using System.IO;
using System.Windows;
using System.Windows.Forms; // 用于文件夹选择器

namespace CadAtlasManager
{
    public partial class CopyFileDialog : Window
    {
        public string ResultFilePath { get; private set; } // 最终生成的完整路径
        private string _originalExt; // 原文件后缀 (.dwg)

        public CopyFileDialog(string sourceFilePath)
        {
            InitializeComponent();

            // 初始化界面数据
            TxtSourceInfo.Text = $"源文件: {Path.GetFileName(sourceFilePath)}";
            _originalExt = Path.GetExtension(sourceFilePath);

            // 默认新文件名：原名_复制
            string fileNameNoExt = Path.GetFileNameWithoutExtension(sourceFilePath);
            TbNewName.Text = $"{fileNameNoExt}_复制";

            // 默认保存路径：当前源文件所在的文件夹
            TbSavePath.Text = Path.GetDirectoryName(sourceFilePath);
        }

        // 浏览按钮：选择保存文件夹
        private void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.SelectedPath = TbSavePath.Text;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    TbSavePath.Text = dialog.SelectedPath;
                }
            }
        }

        // 确定按钮
        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            string newName = TbNewName.Text.Trim();
            string savePath = TbSavePath.Text.Trim();

            if (string.IsNullOrEmpty(newName))
            {
                System.Windows.MessageBox.Show("请输入文件名！");
                return;
            }

            // 组合完整路径
            string fullPath = Path.Combine(savePath, newName + _originalExt);

            // 检查文件是否存在
            if (File.Exists(fullPath))
            {
                var result = System.Windows.MessageBox.Show("目标路径已存在同名文件，是否覆盖？", "覆盖确认", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.No) return;
            }

            ResultFilePath = fullPath;
            this.DialogResult = true; // 关闭窗口并返回 true
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
        }
        // 允许外部设置默认保存路径
        public void SetDefaultSavePath(string path)
        {
            if (System.IO.Directory.Exists(path))
            {
                TbSavePath.Text = path;
            }
        }
    }
}