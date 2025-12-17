namespace CadAtlasManager.Models
{
    // 打印候选对象模型
    public class PlotCandidate
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public bool IsOutdated { get; set; }
        public string StatusText => IsOutdated ? "⚠️ 已修改" : "✅ 未修改";

        // 绑定 CheckBox：如果是过期的，默认选中；如果没过期，默认不选中
        public bool IsSelected { get; set; }

        // 携带最新的指纹数据，方便后续传递
        public string NewTdUpdate { get; set; }
    }
}