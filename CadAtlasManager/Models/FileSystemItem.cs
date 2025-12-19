using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Media;

namespace CadAtlasManager.Models
{
    public enum ExplorerItemType { Folder, File }

    public class FileSystemItem : INotifyPropertyChanged
    {
        public string Name { get; set; }
        public string FullPath { get; set; }
        public ExplorerItemType Type { get; set; }
        public string TypeIcon { get; set; }
        public bool IsRoot { get; set; } = false;
        public FontWeight FontWeight { get; set; } = FontWeights.Normal;

        // --- 新增：属性快捷判断 ---
        public bool IsDirectory => Type == ExplorerItemType.Folder;

        // --- 新增：激活状态（用于文字变红） ---
        private bool _isActive;
        public bool IsActive
        {
            get => _isActive;
            set { if (_isActive != value) { _isActive = value; OnPropertyChanged("IsActive"); } }
        }

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

        private bool _isChecked;
        public bool IsChecked
        {
            get { return _isChecked; }
            set { if (_isChecked != value) { _isChecked = value; OnPropertyChanged("IsChecked"); } }
        }

        public string CreationDate { get; set; }

        private string _versionStatus;
        public string VersionStatus
        {
            get { return _versionStatus; }
            set { if (_versionStatus != value) { _versionStatus = value; OnPropertyChanged("VersionStatus"); } }
        }

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