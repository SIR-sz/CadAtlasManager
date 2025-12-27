using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace CadAtlasManager.Models
{
    public class PlotCandidate : INotifyPropertyChanged
    {
        // 统一使用 FilePath (解决 FullPath 报错)
        public string FilePath { get; set; }

        public string FileName { get; set; }

        private bool _isSelected = true;
        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        private string _status = "待处理";
        public string Status
        {
            get => _status;
            set { _status = value; OnPropertyChanged(); }
        }

        private bool? _isSuccess = null;
        public bool? IsSuccess
        {
            get => _isSuccess;
            set { _isSuccess = value; OnPropertyChanged(); }
        }

        // 补全缺失的属性 (解决 VersionStatus 和 IsOutdated 报错)
        public string VersionStatus { get; set; } = "未知";
        public bool IsOutdated { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}