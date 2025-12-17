using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;

[assembly: CommandClass(typeof(CadAtlasManager.MainPlugin))]

namespace CadAtlasManager
{
    public class MainPlugin
    {
        // 静态变量保持 PaletteSet 实例，保证只有一个侧边栏
        private static PaletteSet _ps = null;
        private static AtlasView _myView = null;

        [CommandMethod("AM")] // 这里定义 CAD 命令为 AM
        public void ShowAtlasManager()
        {
            if (_ps == null)
            {
                // 创建侧边栏容器
                _ps = new PaletteSet("CAD图集管理器");

                _ps.Style = PaletteSetStyles.ShowCloseButton |
                            PaletteSetStyles.ShowPropertiesMenu |
                            PaletteSetStyles.Snappable;

                // 实例化我们的 WPF 界面
                _myView = new AtlasView();

                // 将 WPF 界面放入 ElementHost，再塞入 PaletteSet
                // 注意：AutoCAD 的 PaletteSet 可以直接 Add WPF 控件（2010以后）
                _ps.AddVisual("图集库", _myView);
            }

            // 显示侧边栏
            _ps.Visible = true;
        }
    }
}