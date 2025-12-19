using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;

[assembly: ExtensionApplication(typeof(CadAtlasManager.MainPlugin))] // 必须定义程序集入口
[assembly: CommandClass(typeof(CadAtlasManager.MainPlugin))]

namespace CadAtlasManager
{
    // 继承 IExtensionApplication 接口
    public class MainPlugin : IExtensionApplication
    {
        private static PaletteSet _ps = null;
        private static AtlasView _myView = null;

        // --- IExtensionApplication 接口实现 ---

        public void Initialize()
        {
            // 插件加载时执行的逻辑（可选）
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
                _ps = new PaletteSet("CAD图集管理器");
                _ps.Style = PaletteSetStyles.ShowCloseButton |
                            PaletteSetStyles.ShowPropertiesMenu |
                            PaletteSetStyles.Snappable;

                _myView = new AtlasView();
                _ps.AddVisual("图集库", _myView);
            }
            _ps.Visible = true;
        }
    }
}