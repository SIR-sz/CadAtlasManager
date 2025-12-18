namespace CadAtlasManager.Models
{
    public class PlotCandidate
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public bool IsSelected { get; set; }
        public bool IsOutdated { get; set; }
        public string NewTdUpdate { get; set; }

        // 【新增】用于在列表中显示状态文本 (如 "✅ 最新", "⚠️ 需更新")
        public string VersionStatus { get; set; }
    }
}

//ashcahcaslcla