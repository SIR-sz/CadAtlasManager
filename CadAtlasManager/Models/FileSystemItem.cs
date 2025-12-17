using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;

namespace CadAtlasManager.Models
{
    // 定义枚举
    public enum ExplorerItemType { Folder, File }

    // 文件树节点模型
    public class FileSystemItem : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public string FullPath { get; set; }
        public ExplorerItemType Type { get; set; }
        public string TypeIcon { get; set; }
        public bool IsRoot { get; set; } = false;
        public FontWeight FontWeight { get; set; } = FontWeights.Normal;

        private bool _isExpanded;
        public bool IsExpanded
        {
            get { return _isExpanded; }
            set { if (_isExpanded != value) { _isExpanded = value; OnPropertyChanged("IsExpanded"); } }
        }

        private bool _isItemSelected;
        public bool IsItemSelected
        {
            get { return _isItemSelected; }
            set { if (_isItemSelected != value) { _isItemSelected = value; OnPropertyChanged("IsItemSelected"); } }
        }

        private bool _hasRemark;
        public bool HasRemark
        {
            get { return _hasRemark; }
            set { if (_hasRemark != value) { _hasRemark = value; OnPropertyChanged("HasRemark"); } }
        }

        // 版本状态文字 (例如 "✅", "⚠️", "❓")
        private string _versionStatus;
        public string VersionStatus
        {
            get { return _versionStatus; }
            set { if (_versionStatus != value) { _versionStatus = value; OnPropertyChanged("VersionStatus"); } }
        }

        // 状态颜色 (用于文字变红/变绿)
        private Brush _statusColor = Brushes.Black;
        public Brush StatusColor
        {
            get { return _statusColor; }
            set { if (_statusColor != value) { _statusColor = value; OnPropertyChanged("StatusColor"); } }
        }

        public ObservableCollection<FileSystemItem> Children { get; set; } = new ObservableCollection<FileSystemItem>();

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}