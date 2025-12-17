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

        // --- 新增设置 ---

        // 打印顺序
        public PlotOrderType OrderType { get; set; } = PlotOrderType.Horizontal;

        // 打印比例模式: "Fit" (布满), "1:1", "1:100" 等
        public string ScaleType { get; set; } = "Fit";

        // 自定义比例值 (例如 1:100，则该值为 0.01，或者存分母 100)
        // 为了简化，我们假设 ScaleType 存的是 "1:1" 这种字符串，或者 "Custom"
        // 这里简单起见，我们主要处理 "Fit" 和 "1:1" 以及自定义偏移

        // 打印偏移 (毫米)
        public double OffsetX { get; set; } = 0.0;
        public double OffsetY { get; set; } = 0.0;

        // 是否居中打印 (如果设置了偏移，通常此项为 false)
        public bool PlotCentered { get; set; } = true;
    }
}