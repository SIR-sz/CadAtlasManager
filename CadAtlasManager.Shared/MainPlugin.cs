using Autodesk.AutoCAD.Windows;
using System;
using Autodesk.AutoCAD.Runtime;

// 使用条件编译别名处理 Application 类的位置变化
#if CAD2012
using AcApp = Autodesk.AutoCAD.ApplicationServices.Application;
#else
using AcApp = Autodesk.AutoCAD.ApplicationServices.Core.Application;
#endif

[assembly: ExtensionApplication(typeof(CadAtlasManager.MainPlugin))] // 必须定义程序集入口
[assembly: CommandClass(typeof(CadAtlasManager.MainPlugin))]

namespace CadAtlasManager
{
    // 继承 IExtensionApplication 接口
    public class MainPlugin : IExtensionApplication
    {
        internal static PaletteSet _ps = null;
        private static AtlasView _myView = null;

        // --- IExtensionApplication 接口实现 ---

        public void Initialize()
        {
            // 将所有 Application.DocumentManager 替换为 AcApp.DocumentManager
            var doc = AcApp.DocumentManager.MdiActiveDocument;
        }

        public void Terminate()
        {
            // 插件卸载时执行的逻辑（可选）
        }

        // --- 命令实现 ---

        [CommandMethod("AM")]
        public void ShowAtlasManager()
        {
            if (_ps == null)
            {
                // 修改点 2：增加一个固定的 GUID（标识符）。有了它，AutoCAD 才会自动保存位置。
                // 你可以使用这个生成的 GUID，或者自己生成一个新的。
                Guid paletteGuid = new Guid("A7F3E2B1-4D5E-4B8C-9F0A-1C2B3D4E5F6A");

                _ps = new PaletteSet("CAD图集管理器", paletteGuid); // 传入 GUID
                _ps.Style = PaletteSetStyles.ShowCloseButton |
                            PaletteSetStyles.ShowPropertiesMenu |
                            PaletteSetStyles.Snappable;

                _myView = new AtlasView();
                _ps.AddVisual("图集库", _myView);

                // 修改点 3：从配置中恢复面板的初始大小
                try
                {
                    var config = ConfigManager.Load();
                    if (config.PaletteWidth > 100 && config.PaletteHeight > 100)
                    {
                        // 设置面板的浮动尺寸
                        _ps.Size = new System.Drawing.Size((int)config.PaletteWidth, (int)config.PaletteHeight);
                    }
                }
                catch { }
            }
            _ps.Visible = true;
        }
    }
}