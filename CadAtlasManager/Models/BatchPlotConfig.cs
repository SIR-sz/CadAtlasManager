// [文件: CadAtlasManager/Models/BatchPlotConfig.cs]
namespace CadAtlasManager.Models
{
    public class BatchPlotConfig
    {
        public string PrinterName { get; set; }       // 打印机 (e.g., "DWG To PDF.pc3")
        public string MediaName { get; set; }         // 默认纸张 (e.g., "ISO_A3_(420.00_x_297.00_MM)")
        public string StyleSheet { get; set; }        // 样式表 (e.g., "monochrome.ctb")
        public string TitleBlockNames { get; set; }   // 图框名 (e.g., "TK,A3")
        public bool AutoRotate { get; set; } = true;  // 是否自动旋转
        public bool UseStandardScale { get; set; } = true; // 是否布满图纸
    }
}