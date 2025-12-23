using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using VMProtect;

[assembly: ExtensionApplication(typeof(CadAtlasManager.MainPlugin))]
[assembly: CommandClass(typeof(CadAtlasManager.MainPlugin))]

namespace CadAtlasManager
{
    public class MainPlugin : IExtensionApplication
    {
        internal static PaletteSet _ps = null;
        private static AtlasView _myView = null;

        // =============================================================
        // 改动 1：初始化只负责打招呼，绝不抛异常
        // =============================================================
        public void Initialize()
        {
            // 可以在这里打印一行字，告诉用户插件加载了，输入 ZX 启动
            // Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n[CAD图集管理] 插件已加载，输入 ZX 启动...\n");

            // 即使没授权，也要让它加载成功！否则命令就没了。
        }

        public void Terminate()
        {
        }

        // =============================================================
        // 改动 2：命令启动时才进行“安检”
        // =============================================================
        [CommandMethod("ZX")]
        public void ShowAtlasManager()
        {
            // 1. 先检查授权
            // 如果检查失败（用户没文件，且在弹窗里点了取消），直接 return，不显示界面
            // 但因为没抛异常，插件依然活着，用户可以再次输入 ZX 重试
            if (!CheckLicense())
            {
                return;
            }

            // 2. 授权通过，才显示界面
            if (_ps == null)
            {
                Guid paletteGuid = new Guid("A7F3E2B1-4D5E-4B8C-9F0A-1C2B3D4E5F6A");
                _ps = new PaletteSet("CAD图集项目管理系统", paletteGuid);
                _ps.Style = PaletteSetStyles.ShowCloseButton |
                            PaletteSetStyles.ShowPropertiesMenu |
                            PaletteSetStyles.Snappable;
                _ps.MinimumSize = new System.Drawing.Size(300, 400);

                _myView = new AtlasView();
                _ps.AddVisual("工作台", _myView);

                try
                {
                    var config = ConfigManager.Load();
                    if (config != null && config.PaletteWidth > 100 && config.PaletteHeight > 100)
                    {
                        _ps.Size = new System.Drawing.Size((int)config.PaletteWidth, (int)config.PaletteHeight);
                    }
                }
                catch { }
            }

            _ps.Visible = true;
        }

        // =============================================================
        // 核心验证逻辑 (保持不变，功能最全版)
        // =============================================================
        [VMProtect.Begin]
        private bool CheckLicense()
        {
            // 1. 优先尝试读取本地 key.lic (VMProtect 自动处理)
            int status = (int)SDK.GetSerialNumberState();
            if (status == 0) return true; // 自动通过

            // 2. 补救措施：尝试手动读取并清洗文件 (解决记事本BOM/空格问题)
            try
            {
                string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string licPath = Path.Combine(dllPath, "key.lic");
                if (File.Exists(licPath))
                {
                    string cleanKey = File.ReadAllText(licPath).Trim();
                    if ((int)SDK.SetSerialNumber(cleanKey) == 0)
                    {
                        File.WriteAllText(licPath, cleanKey); // 修正文件
                        return true;
                    }
                }
            }
            catch { }

            // 3. 弹窗处理
            string hwid = SDK.GetCurrentHWID();
            string msg = "CAD图集项目管理系统 - 授权验证失败";
            switch (status)
            {
                case 1: msg = "未检测到有效授权，请激活。"; break;

                // case 2 原来是 "您的授权已被列入黑名单。"
                // ✅ 修改为比较委婉的说法：
                case 2: msg = "当前授权已失效，请联系管理员获取新码。"; break;

                case 3: msg = "您的授权已过期。"; break;
                default: msg = $"未知状态 (代码: {status})，请重新激活。"; break;
            }

            return ShowActivationDialog(msg, hwid);
        }

        private bool ShowActivationDialog(string message, string hwid)
        {
            bool isSuccess = false;
            Form form = new Form();
            form.Text = "软件激活向导";
            form.Size = new System.Drawing.Size(480, 450);
            form.StartPosition = FormStartPosition.CenterScreen;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false; form.MinimizeBox = false;
            form.TopMost = true;

            Label lblMsg = new Label() { Text = message + "\n可以直接输入注册码激活。", Dock = DockStyle.Top, Height = 60, Padding = new Padding(10), ForeColor = System.Drawing.Color.Red, Font = new System.Drawing.Font("微软雅黑", 9F) };

            GroupBox grpHwid = new GroupBox() { Text = "1. 复制机器码发给管理员", Dock = DockStyle.Top, Height = 80, Padding = new Padding(5) };
            TextBox txtHwid = new TextBox() { Text = hwid, Dock = DockStyle.Top, ReadOnly = true, Font = new System.Drawing.Font("Consolas", 10F) };
            Button btnCopy = new Button() { Text = "点击复制机器码", Dock = DockStyle.Bottom, Height = 28 };
            btnCopy.Click += (s, e) => { Clipboard.SetText(hwid); MessageBox.Show("机器码已复制！"); };
            grpHwid.Controls.Add(btnCopy); grpHwid.Controls.Add(txtHwid);

            GroupBox grpInput = new GroupBox() { Text = "2. 输入注册码", Dock = DockStyle.Top, Height = 180, Padding = new Padding(5) };
            TextBox txtKey = new TextBox() { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical, Font = new System.Drawing.Font("Consolas", 9F) };
            grpInput.Controls.Add(txtKey);

            Button btnActivate = new Button() { Text = "立即激活", Dock = DockStyle.Bottom, Height = 50, Font = new System.Drawing.Font("微软雅黑", 10F, System.Drawing.FontStyle.Bold) };

            btnActivate.Click += (s, e) =>
            {
                string keyInput = txtKey.Text.Trim();
                if (string.IsNullOrEmpty(keyInput)) return;

                int res = (int)SDK.SetSerialNumber(keyInput); // 热激活
                if (res == 0)
                {
                    try
                    {
                        string dllPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        File.WriteAllText(Path.Combine(dllPath, "key.lic"), keyInput);
                    }
                    catch { }
                    MessageBox.Show("激活成功！");
                    isSuccess = true;
                    form.Close();
                }
                else
                {
                    MessageBox.Show($"注册码无效 (错误代码: {res})");
                }
            };

            form.Controls.Add(grpInput); form.Controls.Add(grpHwid); form.Controls.Add(lblMsg); form.Controls.Add(btnActivate);
            form.ShowDialog();
            return isSuccess;
        }
    }

}