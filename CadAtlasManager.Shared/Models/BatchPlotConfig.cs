namespace CadAtlasManager.Models
{
    // 定义排序枚举
    public enum PlotOrderType
    {
        Horizontal, // 横向：先左后右，再换行 (Z字形)
        Vertical    // 竖向：先上后下，再换列 (N字形)
    }

    public class BatchPlotConfig
    {
        public string PrinterName { get; set; }
        public string MediaName { get; set; }
        public string StyleSheet { get; set; }
        public string TitleBlockNames { get; set; }
        public bool AutoRotate { get; set; }

        // --- 新增功能：是否强制使用用户选定的纸张（不进行自动匹配） ---
        public bool ForceUseSelectedMedia { get; set; } = false;

        // --- 新增设置 ---

        // 打印顺序
        public PlotOrderType OrderType { get; set; } = PlotOrderType.Horizontal;

        // 打印比例模式: "Fit" (布满), "1:1", "1:100" 等
        public string ScaleType { get; set; } = "Fit";

        // 打印偏移 (毫米)
        public double OffsetX { get; set; } = 0.0;
        public double OffsetY { get; set; } = 0.0;

        // 是否居中打印
        public bool PlotCentered { get; set; } = true;

        // --- 2.0 重构新增：打印选项 (Clean Slate 模式) ---
        // 默认打印线宽，不打印透明度(提升速度)，使用样式表
        public bool PlotWithLineweights { get; set; } = true;
        public bool PlotTransparency { get; set; } = false;
        public bool PlotWithPlotStyles { get; set; } = true;
        public bool HidePaperspaceObjects { get; set; } = false;
    }
}