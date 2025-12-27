using System.Collections.Generic;

namespace CadAtlasManager.Models
{
    // 记录单个 DWG 文件的打印执行情况
    public class PlotFileResult
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public bool IsSuccess { get; set; } // 是否识别并打印成功
        public int PageCount { get; set; }  // 打印出的页数
        public string ErrorMessage { get; set; } // 失败原因（如：未找到图框）
        public List<string> GeneratedPdfs { get; set; } = new List<string>();
    }
}