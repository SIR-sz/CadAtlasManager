using CadAtlasManager.Models; // 确保引用了模型命名空间
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace CadAtlasManager
{
    public class AppConfig
    {
        public List<string> AtlasFolders { get; set; } = new List<string>();
        public List<ProjectItem> Projects { get; set; } = new List<ProjectItem>();
        public string LastActiveProjectPath { get; set; } = "";

        public double PaletteWidth { get; set; } = 350;
        public double PaletteHeight { get; set; } = 700;

        public double ProjectTreeWidth { get; set; } = 200;
        public double ProjectNameColumnWidth { get; set; } = 300;
        public double PlotTreeWidth { get; set; } = 200;
        public double PlotNameColumnWidth { get; set; } = 300;

        // --- 打印相关配置 ---
        public string TitleBlockNames { get; set; } = "TK,A3图框";
        public string LastPrinter { get; set; } = "DWG To PDF.pc3";
        public string LastStyleSheet { get; set; } = "monochrome.ctb";
        public string LastMedia { get; set; } = "";

        // --- 修改3: 新增字段用于记录上一次打印参数 ---
        public PlotOrderType LastOrderType { get; set; } = PlotOrderType.Horizontal;
        public bool LastFitToPaper { get; set; } = false; // 默认不勾选布满
        public string LastScaleType { get; set; } = "1:1"; // 默认 1:1
        public bool LastCenterPlot { get; set; } = true;
        public double LastOffsetX { get; set; } = 0.0;
        public double LastOffsetY { get; set; } = 0.0;
        public bool LastAutoRotate { get; set; } = true;
    }

    public static class ConfigManager
    {
        // 配置文件路径：AppData/Roaming/CadAtlasManager/config.xml
        private static readonly string ConfigPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "CadAtlasManager",
            "config.xml");

        public static AppConfig Load()
        {
            try
            {
                if (!File.Exists(ConfigPath)) return new AppConfig();

                XmlSerializer xs = new XmlSerializer(typeof(AppConfig));
                using (StreamReader sr = new StreamReader(ConfigPath))
                {
                    return (AppConfig)xs.Deserialize(sr);
                }
            }
            catch
            {
                // 加载失败则返回默认配置
                return new AppConfig();
            }
        }

        // =================================================================
        // 直接保存配置对象
        // =================================================================
        public static void Save(AppConfig config)
        {
            try
            {
                string dir = Path.GetDirectoryName(ConfigPath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                XmlSerializer xs = new XmlSerializer(typeof(AppConfig));
                using (StreamWriter sw = new StreamWriter(ConfigPath))
                {
                    xs.Serialize(sw, config);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("保存配置失败: " + ex.Message);
            }
        }

        // =================================================================
        // 【兼容旧代码】只更新列表，不覆盖其他字段
        // =================================================================
        public static void Save(List<string> folders, IEnumerable<ProjectItem> projects, string activePath)
        {
            try
            {
                // 1. 先加载现有配置（关键！这样才能保留 TitleBlockNames 等未传入的字段）
                var config = Load();

                // 2. 更新传入的字段
                config.AtlasFolders = folders ?? new List<string>();
                config.Projects = new List<ProjectItem>(projects);
                config.LastActiveProjectPath = activePath ?? "";

                // 3. 保存完整对象
                Save(config);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("保存列表配置失败: " + ex.Message);
            }
        }
    }
}