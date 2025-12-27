using System.Windows;

namespace CadAtlasManager
{
    public partial class RenameDialog : Window
    {
        public string NewName { get; private set; }

        public RenameDialog(string oldNameWithoutExt)
        {
            InitializeComponent();
            TbNewName.Text = oldNameWithoutExt;
            TbNewName.Focus();
            TbNewName.SelectAll();
        }

        private void BtnOK_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TbNewName.Text))
            {
                MessageBox.Show("名称不能为空");
                return;
            }
            NewName = TbNewName.Text.Trim();
            DialogResult = true;
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}