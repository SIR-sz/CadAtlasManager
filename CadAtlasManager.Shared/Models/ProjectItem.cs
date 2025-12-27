namespace CadAtlasManager.Models
{
    // 项目对象模型
    public class ProjectItem
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public string OutputPath { get; set; }

        public ProjectItem() { }
        public ProjectItem(string name, string path)
        {
            Name = name;
            Path = path;
            // 自动配置输出目录为项目下的 _Plot 文件夹
            OutputPath = System.IO.Path.Combine(path, "_Plot");
        }
    }
}