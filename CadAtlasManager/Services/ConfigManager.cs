using CadAtlasManager.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace CadAtlasManager
{
    // 配置数据结构
    public class AppConfig
    {
        public List<string> AtlasFolders { get; set; } = new List<string>();
        public List<ProjectItem> Projects { get; set; } = new List<ProjectItem>();
        public string LastActiveProjectPath { get; set; } = ""; // 记住上次选中的项目
    }

    public static class ConfigManager
    {
        // 配置文件保存路径：C:\Users\用户名\AppData\Roaming\CadAtlasManager\config.xml
        private static string ConfigPath
        {
            get
            {
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string folder = Path.Combine(appData, "CadAtlasManager");
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                return Path.Combine(folder, "config.xml");
            }
        }

        // 保存配置
        public static void Save(List<string> folders, IEnumerable<ProjectItem> projects, string activeProjectPath)
        {
            try
            {
                var config = new AppConfig
                {
                    AtlasFolders = folders,
                    Projects = new List<ProjectItem>(projects),
                    LastActiveProjectPath = activeProjectPath
                };

                XmlSerializer serializer = new XmlSerializer(typeof(AppConfig));
                using (StreamWriter writer = new StreamWriter(ConfigPath))
                {
                    serializer.Serialize(writer, config);
                }
            }
            catch (Exception ex)
            {
                // 保存失败通常不弹窗打扰用户，除非调试
                System.Diagnostics.Debug.WriteLine("保存配置失败: " + ex.Message);
            }
        }

        // 读取配置
        public static AppConfig Load()
        {
            if (!File.Exists(ConfigPath)) return new AppConfig();

            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(AppConfig));
                using (StreamReader reader = new StreamReader(ConfigPath))
                {
                    return (AppConfig)serializer.Deserialize(reader);
                }
            }
            catch
            {
                return new AppConfig(); // 读取失败则返回默认空配置
            }
        }
    }
}