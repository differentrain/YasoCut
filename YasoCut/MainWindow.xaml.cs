using Microsoft.Win32;

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;

using System.Windows.Input;
using System.Windows.Interop;

using YasoCut.Internals;
using YasoCut.PInvoke;



namespace YasoCut
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private const int WM_HOTKEY = 0x312;
        private const int DWMWA_NCRENDERING_ENABLED = 1;
        private const int DWMWA_NCRENDERING_POLICY = 2;
        private const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;
        private const int NCRP_DISABLED = 1;
        private const int NCRP_ENABLED = 2;

        private static readonly char[] s_invalidChars = Path.GetInvalidPathChars().Concat(new char[] { '<', '>', ':', '"', '/', '\\', '*', '|', '?', '*', ' ', '.' })
                          .Distinct().ToArray();
        private static readonly System.Windows.Forms.FolderBrowserDialog s_folderBrowser = new System.Windows.Forms.FolderBrowserDialog()
        {
            RootFolder = Environment.SpecialFolder.Desktop,
            Description = "选择要保存的目录",
            ShowNewFolderButton = true
        };
        private readonly ushort _atom;
        private bool _inited;
        private bool _isSettingShortCut = false;
        private bool _isShortCutKeyOn = false;
        private WindowInteropHelper _hwndHelper;
        private ModifierKeys _modifierKeys;
        private int _vkey;
        private System.Windows.Forms.NotifyIcon _notifyIcon = null;
        private System.Windows.Forms.ToolStripMenuItem _checkNotExitMenuItem;
        private System.Windows.Forms.ToolStripMenuItem _menuFormatMenuItem;
        private System.Windows.Forms.ToolStripComboBox _comboFormatMenuItem;
        private System.Windows.Forms.ToolStripMenuItem _checkCopyMenuItem;
        private System.Windows.Forms.ToolStripMenuItem _checkCutFullMenuItem;
        private System.Windows.Forms.ToolStripMenuItem _checkShowTipMenuItem;
        private System.Windows.Forms.ToolStripMenuItem _checkClickTipOpenFolderMenuItem;
        private System.Windows.Forms.ToolStripMenuItem _checkClickTipOpenFileMenuItem;
        private System.Windows.Forms.ToolStripMenuItem _checkRemoveAeroMenuItem;
        private System.Windows.Forms.ToolStripMenuItem _checkNotSaveMenuItem;
        private System.Windows.Forms.ToolStripMenuItem _checkTryProtectMenuItem;
        private bool _notExit = true;
        private string _lastImgFile;

        public Mutex MyMutex { get; }

        public MainWindow()
        {
            string[] args = Environment.GetCommandLineArgs();

            if (args.Length >= 2 && args[1] == "-c")
            {
                using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
                {
                    soft.DeleteSubKeyTree("YasoCut");
                }
                MessageBox.Show("所有注册表信息已经清除完毕！", "YasoCut");
                _notExit = false;
                Close();
            }
            else
            {
                string myGuid = Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>().Value;

                MyMutex = new Mutex(true, Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>().Value, out var createNew);
                if (!createNew)
                {
                    _notExit = false;
                    Close();
                }
                else
                {
                    _atom = NativeMethods.GlobalAddAtom(myGuid);
                    InitializeComponent();
                    SetIcon();
                }
            }


        }

        private void SetIcon()
        {
            this._notifyIcon = new System.Windows.Forms.NotifyIcon
            {
                Text = "YasoCut",
                Icon = Properties.Resources.popo,
                Visible = true
            };
            _notifyIcon.BalloonTipClicked += NotifyIcon_BalloonTipClicked;
            _notifyIcon.MouseDoubleClick += NotifyIcon_MouseDoubleClick; ;

            var cms = new System.Windows.Forms.ContextMenuStrip();

            _notifyIcon.ContextMenuStrip = cms;


            _checkNotExitMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "窗口关闭后最小化到托盘",
                CheckOnClick = true,
            };
            _checkNotExitMenuItem.Click += CheckNotExitMenuItem_Click;
            cms.Items.Add(_checkNotExitMenuItem);

            var separatorMenuItem3 = new System.Windows.Forms.ToolStripSeparator
            {
                Text = "-",
            };
            cms.Items.Add(separatorMenuItem3);

            _checkCutFullMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "截图包括阴影区域",
                CheckOnClick = true,
            };
            _checkCutFullMenuItem.Click += CheckCutFullMenuItem_Click;
            cms.Items.Add(_checkCutFullMenuItem);

            _checkTryProtectMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "尝试截取受保护的窗口",
                CheckOnClick = true,
            };
            _checkTryProtectMenuItem.Click += CheckTryProtectMenuItem_Click;
            cms.Items.Add(_checkTryProtectMenuItem);


            _checkRemoveAeroMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "截图移除Aero效果",
                CheckOnClick = true,
            };
            _checkRemoveAeroMenuItem.Click += CheckRemoveAeroMenuItem_Click;
            cms.Items.Add(_checkRemoveAeroMenuItem);


            var separatorMenuItem2 = new System.Windows.Forms.ToolStripSeparator
            {
                Text = "-",
            };
            cms.Items.Add(separatorMenuItem2);

            _checkCopyMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "截图后添加到剪切板",
                CheckOnClick = true,
            };
            _checkCopyMenuItem.Click += CheckCopyMenuItem_Click;
            cms.Items.Add(_checkCopyMenuItem);


            _checkShowTipMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "截图后显示提示",
                CheckOnClick = true,
            };
            _checkShowTipMenuItem.Click += CheckShowTipMenuItem_Click;

            _checkClickTipOpenFolderMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "单击提示打开文件夹",
                CheckOnClick = true,
            };
            _checkClickTipOpenFolderMenuItem.Click += CheckClickTipOpenFolderMenuItem_Click;
            _checkClickTipOpenFileMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "单击提示打开文件",
                CheckOnClick = true,
            };
            _checkClickTipOpenFileMenuItem.Click += CheckClickTipOpenFileMenuItem_Click;
            _checkShowTipMenuItem.DropDown.Items.Add(_checkClickTipOpenFolderMenuItem);
            _checkShowTipMenuItem.DropDown.Items.Add(_checkClickTipOpenFileMenuItem);

            cms.Items.Add(_checkShowTipMenuItem);

            var separatorMenuItem0 = new System.Windows.Forms.ToolStripSeparator
            {
                Text = "-",
            };
            cms.Items.Add(separatorMenuItem0);

            _checkNotSaveMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "不保存到文件",
                CheckOnClick = true,
            };
            _checkNotSaveMenuItem.Click += CheckNotSaveMenuItem_Click;
            cms.Items.Add(_checkNotSaveMenuItem);


            _menuFormatMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            _comboFormatMenuItem = new System.Windows.Forms.ToolStripComboBox
            {
                DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            };
            _menuFormatMenuItem.DropDown.Items.Add(_comboFormatMenuItem);
            cms.Items.Add(_menuFormatMenuItem);
            _comboFormatMenuItem.Items.AddRange(Enum.GetNames(typeof(ImageFormatType)));
            _comboFormatMenuItem.SelectedIndexChanged += ComboFormatMenuItem_SelectedIndexChanged;



            var separatorMenuItem1 = new System.Windows.Forms.ToolStripSeparator
            {
                Text = "-",
            };
            cms.Items.Add(separatorMenuItem1);


            var exitMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "退出",
            };
            exitMenuItem.Click += ExitMenuItem_Click;


            cms.Items.Add(exitMenuItem);

        }

        private void CheckTryProtectMenuItem_Click(object sender, EventArgs e)
        {
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("TryProtect", _checkTryProtectMenuItem.Checked ? 1 : 0, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
        }

        private void CheckNotSaveMenuItem_Click(object sender, EventArgs e)
        {
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("NotSave", _checkNotSaveMenuItem.Checked ? 1 : 0, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
        }

        private void CheckRemoveAeroMenuItem_Click(object sender, EventArgs e)
        {
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("RemoveAero", _checkRemoveAeroMenuItem.Checked ? 1 : 0, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
        }

        private void CheckClickTipOpenFileMenuItem_Click(object sender, EventArgs e)
        {
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("ClickTipAction", _checkClickTipOpenFileMenuItem.Checked ? 2 : 0, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
            if (_checkClickTipOpenFileMenuItem.Checked)
            {
                _checkClickTipOpenFolderMenuItem.Checked = false;
            }
        }

        private void CheckClickTipOpenFolderMenuItem_Click(object sender, EventArgs e)
        {
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("ClickTipAction", _checkClickTipOpenFolderMenuItem.Checked ? 1 : 0, RegistryValueKind.DWord);
                yasocut.Dispose();
            }

            if (_checkClickTipOpenFolderMenuItem.Checked)
            {
                _checkClickTipOpenFileMenuItem.Checked = false;
            }

        }

        private void NotifyIcon_BalloonTipClicked(object sender, EventArgs e)
        {
            if (_lastImgFile == null)
            {
                return;
            }

            if (_checkClickTipOpenFolderMenuItem.Checked)
            {
                ProcessStartInfo psi = new ProcessStartInfo("Explorer.exe")
                {
                    Arguments = $"/e,/select,{_lastImgFile}"
                };
                try
                {
                    Process.Start(psi)?.Dispose();
                }
                catch
                {
                }
            }
            else if (_checkClickTipOpenFileMenuItem.Checked)
            {
                try
                {
                    Process.Start(_lastImgFile)?.Dispose();
                }
                catch
                {
                }
            }
        }

        private void CheckShowTipMenuItem_Click(object sender, EventArgs e)
        {
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("ShowTip", _checkShowTipMenuItem.Checked ? 1 : 0, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
        }

        private void CheckCutFullMenuItem_Click(object sender, EventArgs e)
        {
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("CutWithShadow", _checkCutFullMenuItem.Checked ? 1 : 0, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
        }

        private void CheckCopyMenuItem_Click(object sender, EventArgs e)
        {
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("CopyToClipboard", _checkCopyMenuItem.Checked ? 1 : 0, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
        }



        private void ComboFormatMenuItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!_inited)
            {
                return;
            }
            _menuFormatMenuItem.Text = $"图片格式: {(ImageFormatType)_comboFormatMenuItem.SelectedIndex}";
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("Format", _comboFormatMenuItem.SelectedIndex, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
        }

        private void CheckNotExitMenuItem_Click(object sender, EventArgs e)
        {
            _notExit = _checkNotExitMenuItem.Checked;
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("ExitToTray", _notExit, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
        }

        private void ExitMenuItem_Click(object sender, EventArgs e)
        {
            _notExit = false;
            this.Close();
        }

        private void NotifyIcon_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            this.WindowState = WindowState.Normal;
            this.Show();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (_notExit)
            {
                this.Hide();
                e.Cancel = true;
            }
            else
            {
                if (_atom != 0)
                {
                    NativeMethods.UnregisterHotKey(_hwndHelper.Handle, _atom);
                    NativeMethods.GlobalDeleteAtom(_atom);

                }
                if (_notifyIcon != null)
                {
                    _notifyIcon.Visible = false;
                    _notifyIcon.Dispose();
                }
            }
            base.OnClosing(e);
        }
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            HwndSource source = PresentationSource.FromVisual(this) as HwndSource;
            _hwndHelper = new WindowInteropHelper(this);
            source.AddHook(WndProc);
        }

        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            ReadReg();
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {

            if (msg != WM_HOTKEY)
            {
                return IntPtr.Zero;
            }

            IntPtr handle = NativeMethods.GetForegroundWindow();
            NativeRect rect = new NativeRect();
            Rectangle bounds;
            NativeMethods.SHQueryUserNotificationState(out int winStyle);
            int ncrp = 0;
            bool needRestNcrp = false;
            int affState;
            if (!_checkTryProtectMenuItem.Checked)
            {
                affState = 0;
            }
            else if (!NativeMethods.GetWindowDisplayAffinity(handle, out affState))
            {
                affState = 0;
            }
            else if (affState != 0)
            {
                NativeMethods.SetWindowDisplayAffinity(handle, 0);
            }
            if (winStyle != 2 && winStyle != 3)
            {
                if (_checkRemoveAeroMenuItem.Checked &&
                    NativeMethods.DwmGetWindowAttribute(handle, DWMWA_NCRENDERING_ENABLED, ref ncrp, 4) == 0 &&
                    ncrp == NCRP_ENABLED)
                {
                    ncrp = NCRP_DISABLED;
                    needRestNcrp = NativeMethods.DwmSetWindowAttribute(handle, DWMWA_NCRENDERING_POLICY, ref ncrp, 4) == 0;
                }
                bool workAreaAndTitle = CheckBoxTitle.IsChecked.Value;
                if ((_checkCutFullMenuItem.Checked && workAreaAndTitle) || NativeMethods.DwmGetWindowAttribute(handle, DWMWA_EXTENDED_FRAME_BOUNDS, ref rect, NativeRect.Size) != 0)
                {
                    NativeMethods.GetWindowRect(handle, ref rect);
                }
                if (workAreaAndTitle)
                {
                    bounds = rect.ToRectangle();
                }
                else
                {
                    NativeRect realRect = new NativeRect();
                    NativeMethods.GetClientRect(handle, ref realRect);
                    bounds = rect.ToRectangle(in realRect);
                }
            }
            else
            {
                NativeMethods.GetWindowRect(handle, ref rect);
                bounds = rect.ToRectangle();
            }
            handled = true;
            SaveImage(bounds);

            if (needRestNcrp)
            {
                ncrp = NCRP_ENABLED;
                NativeMethods.DwmSetWindowAttribute(handle, DWMWA_NCRENDERING_POLICY, ref ncrp, 4);
            }
            if (affState != 0)
            {
                NativeMethods.SetWindowDisplayAffinity(handle, affState);
            }

            return IntPtr.Zero;
        }

        private void SaveImage(in Rectangle bounds)
        {
            using (Bitmap bmp = new Bitmap(bounds.Width, bounds.Height))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(new System.Drawing.Point(bounds.Left, bounds.Top), System.Drawing.Point.Empty, bounds.Size);
                }
                var format = (ImageFormatType)_comboFormatMenuItem.SelectedIndex;
                if (_checkNotSaveMenuItem.Checked || _checkCopyMenuItem.Checked)
                {
                    System.Windows.Forms.Clipboard.SetImage(bmp);
                }
                if (!_checkNotSaveMenuItem.Checked)
                {
                    _lastImgFile = $"{TextboxPath.Text}\\{TextboxPrefix.Text}{DateTime.Now:yyyyMMddHHmmssfff}.{format.GetImageExtensionName()}";
                    try
                    {
                        bmp.Save(_lastImgFile, format.GetImageFormat());
                        if (_checkShowTipMenuItem.Checked)
                        {
                            this._notifyIcon.ShowBalloonTip(1000, $"截图成功!\n{_lastImgFile}", "YasoCut", System.Windows.Forms.ToolTipIcon.Info);
                        }
                    }
                    catch (Exception ex)
                    {
                        _lastImgFile = null;
                        this._notifyIcon.ShowBalloonTip(1000, $"截图失败!\n{ex}", "YasoCut", System.Windows.Forms.ToolTipIcon.Info);
                    }
                }

            }
        }

        private void ButtonShotcut_Click(object sender, RoutedEventArgs e)
        {
            if (_isSettingShortCut)
            {
                _isShortCutKeyOn = NativeMethods.RegisterHotKey(_hwndHelper.Handle, _atom, _modifierKeys, _vkey);

                if (!_isShortCutKeyOn)
                {
                    MessageBox.Show("注册全局快捷键失败！", "错误", MessageBoxButton.OK);
                    _modifierKeys = 0;
                    _vkey = 0;
                    TextboxShotcut.Text = string.Empty;
                    WindowTitle.Title = "YasoCut - 快捷键无效";
                    return;
                }
                _isSettingShortCut = false;
                TextboxShotcut.IsEnabled = false;
                TextboxPrefix.IsEnabled = true;
                TextboxPath.IsEnabled = true;
                ButtonOpen.IsEnabled = true;
                ButtonSelect.IsEnabled = true;
                CheckBoxTitle.IsEnabled = true;
                ButtonShotcut.Content = "设置快捷键";
                WindowTitle.Title = "YasoCut";
                using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
                {
                    long shortcut = ((long)(_modifierKeys) << 32)
                                    | (long)_vkey;
                    RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                    yasocut.SetValue("Shortcut", shortcut, RegistryValueKind.QWord);
                }
            }
            else
            {
                ButtonShotcut.Content = "保存快捷键";
                NativeMethods.UnregisterHotKey(_hwndHelper.Handle, _atom);
                ButtonShotcut.IsEnabled = false;
                _isShortCutKeyOn = false;
                _isSettingShortCut = true;
                TextboxShotcut.IsEnabled = true;
                TextboxShotcut.Text = string.Empty;
                TextboxPrefix.IsEnabled = false;
                TextboxPath.IsEnabled = false;
                ButtonOpen.IsEnabled = false;
                ButtonSelect.IsEnabled = false;
                CheckBoxTitle.IsEnabled = false;
                TextboxShotcut.Focus();
            }

        }

        private void TextboxShotcut_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.LWin:
                case Key.RWin:
                case Key.LeftShift:
                case Key.RightShift:
                case Key.LeftCtrl:
                case Key.RightCtrl:
                case Key.LeftAlt:
                case Key.RightAlt:
                case Key.F12:
                case Key.System:
                    TextboxShotcut.Text = String.Empty;
                    _modifierKeys = ModifierKeys.None;
                    ButtonShotcut.IsEnabled = false;
                    return;
                default:
                    break;
            }

            if (Keyboard.IsKeyDown(Key.F12) ||
                Keyboard.IsKeyDown(Key.System) ||
                ((Keyboard.Modifiers & ModifierKeys.Alt) == ModifierKeys.None &&
                 (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.None &&
                 (Keyboard.Modifiers & ModifierKeys.Windows) == ModifierKeys.None &&
                 (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.None))
            {
                TextboxShotcut.Text = String.Empty;
                _modifierKeys = ModifierKeys.None;
                ButtonShotcut.IsEnabled = false;
                return;
            }

            if (e.Key == Key.Tab)
            {
                e.Handled = true;
            }
            _modifierKeys = Keyboard.Modifiers;
            ButtonShotcut.IsEnabled = true;
            _vkey = KeyInterop.VirtualKeyFromKey(e.Key);
            TextboxShotcut.Text = $"{_modifierKeys}, {e.Key}";
        }

        private void ReadReg()
        {
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);

                if (!(yasocut.GetValue("Path") is string path) || !Directory.Exists(path))
                {
                    path = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                    yasocut.SetValue("Path", path, RegistryValueKind.String);
                }
                TextboxPath.Text = path;

                if (!(yasocut.GetValue("Prefix") is string prefix))
                {
                    prefix = string.Empty;
                    yasocut.SetValue("Prefix", prefix, RegistryValueKind.String);
                }
                TextboxPrefix.Text = prefix;

                if (!(yasocut.GetValue("IncludeTitle") is int includeTitle))
                {
                    includeTitle = 1;
                    yasocut.SetValue("IncludeTitle", includeTitle, RegistryValueKind.DWord);
                }
                CheckBoxTitle.IsChecked = includeTitle != 0;

                if (!(yasocut.GetValue("TryProtect") is int tryProtect))
                {
                    tryProtect = 0;
                    yasocut.SetValue("TryProtect", tryProtect, RegistryValueKind.DWord);
                }
                _checkTryProtectMenuItem.Checked = tryProtect != 0;

                if (!(yasocut.GetValue("ExitToTray") is int notClose))
                {
                    notClose = 1;
                    yasocut.SetValue("ExitToTray", notClose, RegistryValueKind.DWord);
                }
                _checkNotExitMenuItem.Checked = _notExit = notClose != 0;

                if (!(yasocut.GetValue("CopyToClipboard") is int copyToCB))
                {
                    copyToCB = 0;
                    yasocut.SetValue("CopyToClipboard", copyToCB, RegistryValueKind.DWord);
                }
                _checkCopyMenuItem.Checked = copyToCB != 0;

                if (!(yasocut.GetValue("CutWithShadow") is int cutWithShadow))
                {
                    cutWithShadow = 0;
                    yasocut.SetValue("CutWithShadow", cutWithShadow, RegistryValueKind.DWord);
                }
                _checkCutFullMenuItem.Checked = cutWithShadow != 0;

                if (!(yasocut.GetValue("RemoveAero") is int removeAero))
                {
                    removeAero = 0;
                    yasocut.SetValue("RemoveAero", removeAero, RegistryValueKind.DWord);
                }
                _checkRemoveAeroMenuItem.Checked = removeAero != 0;

                if (!(yasocut.GetValue("ShowTip") is int showTip))
                {
                    showTip = 0;
                    yasocut.SetValue("ShowTip", showTip, RegistryValueKind.DWord);
                }
                _checkShowTipMenuItem.Checked = showTip != 0;

                if (!(yasocut.GetValue("NotSave") is int notSave))
                {
                    notSave = 0;
                    yasocut.SetValue("NotSave", notSave, RegistryValueKind.DWord);
                }
                _checkNotSaveMenuItem.Checked = notSave != 0;

                if (!(yasocut.GetValue("ClickTipAction") is int clickTipAction))
                {
                    clickTipAction = 0;
                    yasocut.SetValue("ClickTipAction", clickTipAction, RegistryValueKind.DWord);
                }
                switch (clickTipAction)
                {
                    case 1:
                        _checkClickTipOpenFolderMenuItem.Checked = true;
                        _checkClickTipOpenFileMenuItem.Checked = false;
                        break;
                    case 2:
                        _checkClickTipOpenFolderMenuItem.Checked = false;
                        _checkClickTipOpenFileMenuItem.Checked = true;
                        break;
                    default:
                        _checkClickTipOpenFolderMenuItem.Checked = false;
                        _checkClickTipOpenFileMenuItem.Checked = false;
                        break;
                }

                if (!(yasocut.GetValue("Format") is int format))
                {
                    format = 0;
                    yasocut.SetValue("Format", format, RegistryValueKind.DWord);
                }
                _comboFormatMenuItem.SelectedIndex = format;
                _menuFormatMenuItem.Text = $"图片格式: {(ImageFormatType)_comboFormatMenuItem.SelectedIndex}";

                if (!(yasocut.GetValue("Shortcut") is long shortcut) || shortcut == 0)
                {
                    shortcut =
                        ((long)(ModifierKeys.Control | ModifierKeys.Alt | ModifierKeys.Shift) << 32)
                        | (long)KeyInterop.VirtualKeyFromKey(Key.P);
                    yasocut.SetValue("Shortcut", shortcut, RegistryValueKind.QWord);
                }
                _modifierKeys = (ModifierKeys)(shortcut >> 32);
                _vkey = (int)(shortcut & 0x7FFFFFFF);
                _isShortCutKeyOn = NativeMethods.RegisterHotKey(_hwndHelper.Handle, _atom, _modifierKeys, _vkey);
                if (_isShortCutKeyOn)
                {
                    TextboxShotcut.Text = $"{_modifierKeys}, {KeyInterop.KeyFromVirtualKey(_vkey)}";
                    WindowTitle.Title = "YasoCut";
                }
                else
                {
                    MessageBox.Show("注册全局快捷键失败！", "错误", MessageBoxButton.OK);
                    _modifierKeys = 0;
                    _vkey = 0;
                    TextboxShotcut.Text = string.Empty;
                    yasocut.SetValue("Shortcut", 0, RegistryValueKind.QWord);
                    WindowTitle.Title = "YasoCut - 快捷键无效";
                }
                yasocut.Dispose();
            }
            _inited = true;
        }

        private void CheckBoxTitle_Checked(object sender, RoutedEventArgs e)
        {
            if (!_inited)
            {
                return;
            }

            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("IncludeTitle", CheckBoxTitle.IsChecked.Value ? 1 : 0, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
        }

        private void ButtonOpen_Click(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(TextboxPath.Text))
            {
                using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
                {
                    RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                    string path = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
                    yasocut.SetValue("Path", path, RegistryValueKind.String);
                    TextboxPath.Text = path;
                    yasocut.Dispose();
                }
            }
            try
            {
                Process.Start(TextboxPath.Text)?.Dispose();
            }
            catch
            {
            }
        }

        private void ButtonSelect_Click(object sender, RoutedEventArgs e)
        {
            if (s_folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
                {
                    RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                    yasocut.SetValue("Path", s_folderBrowser.SelectedPath, RegistryValueKind.String);
                    TextboxPath.Text = s_folderBrowser.SelectedPath;
                    yasocut.Dispose();
                }
            }
        }

        private void TextboxPrefix_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_inited)
            {
                return;
            }

            string prefix = TextboxPrefix.Text;
            if (IsInvalid(prefix))
            {
                MessageBox.Show("前缀中包含非法字符。", "错误", MessageBoxButton.OK);
                TextboxPrefix.Text = string.Empty;
                return;
            }

            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue("Prefix", prefix, RegistryValueKind.String);
                yasocut.Dispose();
            }
        }


        private static bool IsInvalid(string prefix)
        {
            if (string.IsNullOrEmpty(prefix))
            {
                return false;
            }

            prefix = prefix.ToUpper();

            if (prefix.StartsWith("COM") ||
                prefix.StartsWith("LPT") ||
                prefix[0] == '.' ||
                prefix[0] == '_' ||
                prefix[0] == '-')
            {
                return true;
            }

            foreach (var chPrefix in prefix)
            {
                foreach (var chInvalid in s_invalidChars)
                {
                    if (chPrefix == chInvalid)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

    }
}
