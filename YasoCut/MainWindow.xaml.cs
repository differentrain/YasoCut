using Microsoft.Win32;

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Threading.Tasks;
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
        private System.Windows.Forms.ToolStripMenuItem _checkSilentMenuItem;
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
        private System.Windows.Forms.ToolStripMenuItem _checkComMsMenuItem;
        private System.Windows.Forms.ToolStripMenuItem _checkObModeMenuItem;
        private bool _notExit = true;
        private string _lastImgFile;
        private Thread _backThread;
        private bool _silentMode = false;
        //   private Thread _backThread2;
        //    private readonly Queue<Bitmap> _queue = new Queue<Bitmap>();

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
                MessageBox.Show("所有注册表信息已经清除完毕！", "YasoCut", MessageBoxButton.OK, MessageBoxImage.Information);
                _notExit = false;
                Close();
            }
            else
            {
                string myGuid = Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>().Value;

                MyMutex = new Mutex(true, Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>().Value, out var createNew);
#if DEBUG
                createNew = true;
#endif
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

            _checkSilentMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "禁用所有气泡提示",
                CheckOnClick = true,
            };
            _checkSilentMenuItem.Click += CheckSilentMenuItem_Click; ;
            cms.Items.Add(_checkSilentMenuItem);

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

            _checkComMsMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "连续模式",
                CheckOnClick = true,
            };
            _checkComMsMenuItem.Click += CheckComMsMenuItem_Click;
            cms.Items.Add(_checkComMsMenuItem);

            _checkObModeMenuItem = new System.Windows.Forms.ToolStripMenuItem
            {
                Text = "爆发模式",
                CheckOnClick = true,
            };
            _checkObModeMenuItem.Click += CheckObModeMenuItem_Click;
            cms.Items.Add(_checkObModeMenuItem);

            var separatorMenuItem4 = new System.Windows.Forms.ToolStripSeparator
            {
                Text = "-",
            };
            cms.Items.Add(separatorMenuItem4);

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

        private void CheckSilentMenuItem_Click(object sender, EventArgs e)
        {
            _silentMode = _checkSilentMenuItem.Checked;
        }



        private void CheckObModeMenuItem_Click(object sender, EventArgs e)
        {
            if (_checkObModeMenuItem.Checked)
            {
                _notifyIcon.Text = WindowTitle.Title = "YasoCut - 爆发模式";
                _checkComMsMenuItem.Checked = false;
            }
            else
            {
                _notifyIcon.Text = WindowTitle.Title = "YasoCut";
            }


        }

        private void CheckComMsMenuItem_Click(object sender, EventArgs e)
        {
            if (_checkComMsMenuItem.Checked)
            {
                _notifyIcon.Text = WindowTitle.Title = "YasoCut - 连续模式";
                _checkObModeMenuItem.Checked = false;
            }
            else
            {
                _notifyIcon.Text = WindowTitle.Title = "YasoCut";
            }
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

                if (_isRunning != 0)
                {
                    _isRunning = 0;
                    while (_backThread.IsAlive /*|| _backThread2.IsAlive*/)
                    {
                        Thread.Sleep(1);
                    }
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

        private int _isRunning = 0;

        private void ShowTips(string title, string text, System.Windows.Forms.ToolTipIcon icon)
        {
            if (_silentMode)
            {
                return;
            }
            this._notifyIcon.ShowBalloonTip(1000, title, text, icon);
        }
        private void BackWorkThreadCon()
        {
            Stopwatch sw = new Stopwatch();
            IntPtr lastHandle = IntPtr.Zero;
            NativeRect rect = new NativeRect();
            NativeRect realRect = new NativeRect();
            Rectangle bounds;
            IntPtr handle;
            bool isFullScreen = false;
            int ncrp = 0;
            bool needRestNcrp = false;
            int affState = 0;
            long lastTick;
            long con;
            var format = (ImageFormatType)_comboFormatMenuItem.SelectedIndex;
            string extension = format.GetImageExtensionName();
            ImageFormat formatR = format.GetImageFormat();
            bool workAreaAndTitle = false;
            bool tryProtect = _checkTryProtectMenuItem.Checked;
            bool cutFull = _checkCutFullMenuItem.Checked;
            bool removeAero = _checkRemoveAeroMenuItem.Checked;
            //  bool isShowTip = _checkShowTipMenuItem.Checked;
            string v = null;
            string prefix = null;
            string folder = null;
            _lastImgFile = null;
            Bitmap bmp = null;
            Graphics g = null;
            Dispatcher.Invoke(() =>
            {
                workAreaAndTitle = CheckBoxTitle.IsChecked.Value;
                v = TextboxMs.Text;
                prefix = TextboxPrefix.Text;
                folder = TextboxPath.Text;
            }, System.Windows.Threading.DispatcherPriority.Send);

            while (folder == null)
            {
                Thread.Sleep(1);
            }
            if (string.IsNullOrEmpty(v))
            {
                con = 1;
            }
            else
            {
                con = int.Parse(v);
            }
            sw.Start();
            while (_isRunning != 0)
            {
                lastTick = sw.ElapsedMilliseconds;
                handle = NativeMethods.GetForegroundWindow();
                if (handle != lastHandle)
                {
                    if (lastHandle != IntPtr.Zero)
                    {

                        if (needRestNcrp)
                        {
                            ncrp = NCRP_ENABLED;
                            NativeMethods.DwmSetWindowAttribute(handle, DWMWA_NCRENDERING_POLICY, ref ncrp, 4);
                        }
                        if (affState != 0)
                        {
                            NativeMethods.SetWindowDisplayAffinity(handle, affState);
                        }
                    }
                    lastHandle = handle;
                    NativeMethods.SHQueryUserNotificationState(out int winStyle);
                    isFullScreen = winStyle == 2 || winStyle == 3;
                    ncrp = 0;
                    needRestNcrp = false;
                    if (!tryProtect)
                    {
                        affState = 0;
                    }
                    else if (!NativeMethods.GetWindowDisplayAffinity(lastHandle, out affState))
                    {
                        affState = 0;
                    }
                    else if (affState != 0)
                    {
                        NativeMethods.SetWindowDisplayAffinity(lastHandle, 0);
                    }
                    if (!isFullScreen)
                    {
                        if (removeAero &&
                             NativeMethods.DwmGetWindowAttribute(lastHandle, DWMWA_NCRENDERING_ENABLED, ref ncrp, 4) == 0 &&
                             ncrp == NCRP_ENABLED)
                        {
                            ncrp = NCRP_DISABLED;
                            needRestNcrp = NativeMethods.DwmSetWindowAttribute(lastHandle, DWMWA_NCRENDERING_POLICY, ref ncrp, 4) == 0;
                        }
                    }
                }
                if (!isFullScreen)
                {
                    //if (workAreaAndTitle)
                    //{
                    //    if (cutFull || NativeMethods.DwmGetWindowAttribute(lastHandle, DWMWA_EXTENDED_FRAME_BOUNDS, ref rect, NativeRect.Size) != 0)
                    //    {
                    //        NativeMethods.GetWindowRect(lastHandle, ref rect);
                    //    }
                    //    bounds = rect.ToRectangle();
                    //}
                    //else
                    //{
                    //    if (NativeMethods.DwmGetWindowAttribute(lastHandle, DWMWA_EXTENDED_FRAME_BOUNDS, ref rect, NativeRect.Size) != 0)
                    //    {
                    //        NativeMethods.GetWindowRect(lastHandle, ref rect);
                    //    }
                    //    NativeMethods.GetClientRect(lastHandle, ref realRect);
                    //    bounds = rect.ToRectangle(in realRect);
                    //}
                    if ((cutFull && workAreaAndTitle) || NativeMethods.DwmGetWindowAttribute(lastHandle, DWMWA_EXTENDED_FRAME_BOUNDS, ref rect, NativeRect.Size) != 0)
                    {
                        NativeMethods.GetWindowRect(lastHandle, ref rect);
                    }
                    if (workAreaAndTitle)
                    {
                        bounds = rect.ToRectangle();
                    }
                    else
                    {

                        NativeMethods.GetClientRect(lastHandle, ref realRect);
                        bounds = rect.ToRectangle(in realRect);
                    }
                }
                else
                {
                    NativeMethods.GetWindowRect(lastHandle, ref rect);
                    bounds = rect.ToRectangle();
                }

                try
                {
                    bmp = new Bitmap(bounds.Width, bounds.Height);
                    g = Graphics.FromImage(bmp);
                    g.CopyFromScreen(new System.Drawing.Point(bounds.Left, bounds.Top), System.Drawing.Point.Empty, bounds.Size);
                }
                catch (Exception ex)
                {
                    ShowTips($"截图失败", $"\n{ex}", System.Windows.Forms.ToolTipIcon.Error);
                    if (bmp != null)
                    {
                        bmp.Dispose();
                        bmp = null;
                        continue;
                    }
                }
                finally
                {
                    g?.Dispose();
                }
                try
                {
                    bmp.Save($"{folder}\\{prefix}{DateTime.Now:yyyyMMddHHmmssfff}_con.{extension}", formatR);
                }
                catch (Exception ex)
                {
                    ShowTips($"截图保存失败", $"\n{ex}", System.Windows.Forms.ToolTipIcon.Error);
                    break;
                }

                bmp.Dispose();
                lastTick = sw.ElapsedMilliseconds - lastTick;
                if (lastTick < con)
                {
                    Thread.Sleep((int)(con - lastTick));
                }
            }
            sw.Stop();
        }

        private void BackWorkThreadOB()
        {
            Stopwatch sw = new Stopwatch();
            IntPtr lastHandle = IntPtr.Zero;
            NativeRect rect = new NativeRect();
            NativeRect realRect = new NativeRect();
            Rectangle bounds;
            IntPtr handle;
            bool isFullScreen = false;
            int ncrp = 0;
            bool needRestNcrp = false;
            int affState = 0;
            long lastTick;
            int con, count;
            var format = (ImageFormatType)_comboFormatMenuItem.SelectedIndex;
            string extension = format.GetImageExtensionName();
            ImageFormat formatR = format.GetImageFormat();
            bool workAreaAndTitle = false;
            bool tryProtect = _checkTryProtectMenuItem.Checked;
            bool cutFull = _checkCutFullMenuItem.Checked;
            bool removeAero = _checkRemoveAeroMenuItem.Checked;
            //  bool isShowTip = _checkShowTipMenuItem.Checked;
            string vMs = null;
            string prefix = null;
            string folder = null;
            _lastImgFile = null;
            Bitmap bmp = null;
            Graphics g = null;
            string vCount = null;
            Dispatcher.Invoke(() =>
            {
                workAreaAndTitle = CheckBoxTitle.IsChecked.Value;
                vMs = TextboxMs.Text;
                prefix = TextboxPrefix.Text;
                folder = TextboxPath.Text;
                vCount = TextboxOB.Text;
            }, System.Windows.Threading.DispatcherPriority.Send);

            while (vCount == null)
            {
                Thread.Sleep(1);
            }
            if (string.IsNullOrEmpty(vMs) || !int.TryParse(vMs, out con))
            {
                con = 1;
            }

            if (string.IsNullOrEmpty(vCount) || !int.TryParse(vCount, out count))
            {
                count = 1;
            }
            sw.Start();
            int i = 0;
            while (_isRunning != 0 && i < count)
            {
                lastTick = sw.ElapsedMilliseconds;
                handle = NativeMethods.GetForegroundWindow();
                if (handle != lastHandle)
                {
                    if (lastHandle != IntPtr.Zero)
                    {

                        if (needRestNcrp)
                        {
                            ncrp = NCRP_ENABLED;
                            NativeMethods.DwmSetWindowAttribute(handle, DWMWA_NCRENDERING_POLICY, ref ncrp, 4);
                        }
                        if (affState != 0)
                        {
                            NativeMethods.SetWindowDisplayAffinity(handle, affState);
                        }
                    }
                    lastHandle = handle;
                    NativeMethods.SHQueryUserNotificationState(out int winStyle);
                    isFullScreen = winStyle == 2 || winStyle == 3;
                    ncrp = 0;
                    needRestNcrp = false;
                    if (!tryProtect)
                    {
                        affState = 0;
                    }
                    else if (!NativeMethods.GetWindowDisplayAffinity(lastHandle, out affState))
                    {
                        affState = 0;
                    }
                    else if (affState != 0)
                    {
                        NativeMethods.SetWindowDisplayAffinity(lastHandle, 0);
                    }
                    if (!isFullScreen)
                    {
                        if (removeAero &&
                             NativeMethods.DwmGetWindowAttribute(lastHandle, DWMWA_NCRENDERING_ENABLED, ref ncrp, 4) == 0 &&
                             ncrp == NCRP_ENABLED)
                        {
                            ncrp = NCRP_DISABLED;
                            needRestNcrp = NativeMethods.DwmSetWindowAttribute(lastHandle, DWMWA_NCRENDERING_POLICY, ref ncrp, 4) == 0;
                        }
                    }
                }
                if (!isFullScreen)
                {
                    //if (workAreaAndTitle)
                    //{
                    //    if (cutFull || NativeMethods.DwmGetWindowAttribute(lastHandle, DWMWA_EXTENDED_FRAME_BOUNDS, ref rect, NativeRect.Size) != 0)
                    //    {
                    //        NativeMethods.GetWindowRect(lastHandle, ref rect);
                    //    }
                    //    bounds = rect.ToRectangle();
                    //}
                    //else
                    //{
                    //    if (NativeMethods.DwmGetWindowAttribute(lastHandle, DWMWA_EXTENDED_FRAME_BOUNDS, ref rect, NativeRect.Size) != 0)
                    //    {
                    //        NativeMethods.GetWindowRect(lastHandle, ref rect);
                    //    }
                    //    NativeMethods.GetClientRect(lastHandle, ref realRect);
                    //    bounds = rect.ToRectangle(in realRect);
                    //}
                    if ((cutFull && workAreaAndTitle) || NativeMethods.DwmGetWindowAttribute(lastHandle, DWMWA_EXTENDED_FRAME_BOUNDS, ref rect, NativeRect.Size) != 0)
                    {
                        NativeMethods.GetWindowRect(lastHandle, ref rect);
                    }
                    if (workAreaAndTitle)
                    {
                        bounds = rect.ToRectangle();
                    }
                    else
                    {

                        NativeMethods.GetClientRect(lastHandle, ref realRect);
                        bounds = rect.ToRectangle(in realRect);
                    }
                }
                else
                {
                    NativeMethods.GetWindowRect(lastHandle, ref rect);
                    bounds = rect.ToRectangle();
                }

                try
                {
                    bmp = new Bitmap(bounds.Width, bounds.Height);
                    g = Graphics.FromImage(bmp);
                    g.CopyFromScreen(new System.Drawing.Point(bounds.Left, bounds.Top), System.Drawing.Point.Empty, bounds.Size);
                }
                catch (Exception ex)
                {
                    ShowTips($"截图失败", $"\n{ex}", System.Windows.Forms.ToolTipIcon.Error);
                    if (bmp != null)
                    {
                        bmp.Dispose();
                        bmp = null;
                        continue;
                    }
                }
                finally
                {
                    g?.Dispose();
                }

                try
                {
                    bmp.Save($"{folder}\\{prefix}{DateTime.Now:yyyyMMddHHmmssfff}_ob.{extension}", formatR);
                    ++i;
                }
                catch (Exception ex)
                {
                    ShowTips($"截图保存失败", $"\n{ex}", System.Windows.Forms.ToolTipIcon.Error);
                    break;
                }

                bmp.Dispose();
                lastTick = sw.ElapsedMilliseconds - lastTick;
                if (lastTick < con)
                {
                    Thread.Sleep((int)(con - lastTick));
                }
            }

            if (Interlocked.CompareExchange(ref _isRunning, 0, 1) != 0)
            {
                this.Dispatcher.Invoke(() =>
                {
                    GridMain.IsEnabled = true;
                    _menuFormatMenuItem.Enabled = true;
                    _comboFormatMenuItem.Enabled = true;
                    _checkCopyMenuItem.Enabled = true;
                    _checkCutFullMenuItem.Enabled = true;
                    _checkShowTipMenuItem.Enabled = true;
                    _checkClickTipOpenFolderMenuItem.Enabled = true;
                    _checkClickTipOpenFileMenuItem.Enabled = true;
                    _checkRemoveAeroMenuItem.Enabled = true;
                    _checkNotSaveMenuItem.Enabled = true;
                    _checkTryProtectMenuItem.Enabled = true;
                    _checkComMsMenuItem.Enabled = true;
                    _checkObModeMenuItem.Enabled = true;
                    _notifyIcon.Text = WindowTitle.Title = $"YasoCut - 爆发模式";
                    ShowTips($"YasoCut", $"停止爆发截图", System.Windows.Forms.ToolTipIcon.Info);
                });
            }
            sw.Stop();
        }

        //  private static readonly object s_locker = new object();


        //private void SaveImageConCore(string extension, ImageFormat formatR)
        //{

        //    Bitmap bmp = null;
        //    bool isShowTip = _checkShowTipMenuItem.Checked;
        //    string fileName;
        //    while (_isRunning || _queue.Count > 0)
        //    {
        //        if (_queue.Count > 0)
        //        {
        //            lock (s_locker)
        //            {
        //                bmp = _queue.Dequeue();
        //            }
        //            fileName = $"{_prefix}{DateTime.Now:yyyyMMddHHmmssfff}.{extension}";
        //            _lastImgFile = $"{_folder}\\{fileName}";
        //            try
        //            {
        //                bmp.Save(_lastImgFile, formatR);
        //                if (isShowTip)
        //                {
        //                    this._notifyIcon.ShowBalloonTip(1000, "截图保存成功", $"{_folder}\n{fileName}", System.Windows.Forms.ToolTipIcon.Info);
        //                }
        //            }
        //            catch (Exception ex)
        //            {
        //                _lastImgFile = null;
        //                this._notifyIcon.ShowBalloonTip(1000, $"截图保存失败", $"\n{ex}", System.Windows.Forms.ToolTipIcon.Info);
        //            }

        //            bmp.Dispose();
        //        }
        //        Thread.Sleep(20);
        //    }
        //}

        private void SwitchConOrOb(ThreadStart ts, string type)
        {

            bool b = _isRunning != 0;


            GridMain.IsEnabled = b;
            _menuFormatMenuItem.Enabled = b;
            _comboFormatMenuItem.Enabled = b;
            _checkCopyMenuItem.Enabled = b;
            _checkCutFullMenuItem.Enabled = b;
            _checkShowTipMenuItem.Enabled = b;
            _checkClickTipOpenFolderMenuItem.Enabled = b;
            _checkClickTipOpenFileMenuItem.Enabled = b;
            _checkRemoveAeroMenuItem.Enabled = b;
            _checkNotSaveMenuItem.Enabled = b;
            _checkTryProtectMenuItem.Enabled = b;
            _checkComMsMenuItem.Enabled = b;
            _checkObModeMenuItem.Enabled = b;

            if (b)
            {
                if (Interlocked.CompareExchange(ref _isRunning, 0, 1) != 0)
                {
                    _notifyIcon.Text = WindowTitle.Title = $"YasoCut - {type}模式";
                    while (_backThread.IsAlive /*|| _backThread2.IsAlive*/)
                    {
                        Thread.Sleep(1);
                    }
                    _backThread = null;
                    ShowTips($"YasoCut", $"停止{type}截图", System.Windows.Forms.ToolTipIcon.Info);
                }
            }
            else
            {
                _isRunning = 1;
                _notifyIcon.Text = WindowTitle.Title = $"YasoCut - {type}模式(执行中)";

                ShowTips($"YasoCut", $"启动{type}截图", System.Windows.Forms.ToolTipIcon.Info);
                _backThread = new Thread(ts)
                {
                    Priority = ThreadPriority.Highest,
                    IsBackground = true
                };
                _backThread.Start();
              
            }
        }


        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {

            if (msg != WM_HOTKEY)
            {
                return IntPtr.Zero;
            }

            if (_checkComMsMenuItem.Checked)
            {
                SwitchConOrOb(BackWorkThreadCon, "连续");
            }
            else if (_checkObModeMenuItem.Checked)
            {
                SwitchConOrOb(BackWorkThreadOB, "爆发");
            }
            else
            {
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

                SaveImage(bounds, TextboxPrefix.Text, TextboxPath.Text);
                if (needRestNcrp)
                {
                    ncrp = NCRP_ENABLED;
                    NativeMethods.DwmSetWindowAttribute(handle, DWMWA_NCRENDERING_POLICY, ref ncrp, 4);
                }
                if (affState != 0)
                {
                    NativeMethods.SetWindowDisplayAffinity(handle, affState);
                }
            }
            handled = true;
            return IntPtr.Zero;
        }

        private void SaveImage(in Rectangle bounds, string prefix, string folder)
        {
            bool notSave = _checkNotSaveMenuItem.Checked;
            Bitmap bmp = null;
            Graphics g = null;
            try
            {
                bmp = new Bitmap(bounds.Width, bounds.Height);
                g = Graphics.FromImage(bmp);
                g.CopyFromScreen(new System.Drawing.Point(bounds.Left, bounds.Top), System.Drawing.Point.Empty, bounds.Size);
            }
            catch (Exception ex)
            {
                ShowTips($"截图失败", $"\n{ex}", System.Windows.Forms.ToolTipIcon.Error);
                if (bmp != null)
                {
                    bmp.Dispose();
                    bmp = null;
                    return;
                }
            }
            finally
            {
                g?.Dispose();
            }

            var format = (ImageFormatType)_comboFormatMenuItem.SelectedIndex;

            if (notSave || _checkCopyMenuItem.Checked)
            {
                try
                {
                    System.Windows.Forms.Clipboard.SetImage(bmp);
                }
                catch (Exception ex)
                {
                    ShowTips($"截图到剪切板失败", $"\n{ex}", System.Windows.Forms.ToolTipIcon.Error);
                }
            }

            if (!notSave)
            {
                var fileName = $"{prefix}{DateTime.Now:yyyyMMddHHmmssfff}.{format.GetImageExtensionName()}";
                _lastImgFile = $"{folder}\\{fileName}";
                try
                {
                    bmp.Save(_lastImgFile, format.GetImageFormat());
                    if (_checkShowTipMenuItem.Checked)
                    {
                        ShowTips("截图保存成功", $"{folder}\n{fileName}", System.Windows.Forms.ToolTipIcon.Info);
                    }
                }
                catch (Exception ex)
                {
                    _lastImgFile = null;
                    ShowTips($"截图保存失败", $"\n{ex}", System.Windows.Forms.ToolTipIcon.Error);
                }
            }
            bmp.Dispose();
        }

        private void ButtonShotcut_Click(object sender, RoutedEventArgs e)
        {
            if (_isSettingShortCut)
            {
                _isShortCutKeyOn = NativeMethods.RegisterHotKey(_hwndHelper.Handle, _atom, _modifierKeys, _vkey);

                if (!_isShortCutKeyOn)
                {
                    MessageBox.Show("注册全局快捷键失败！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    _modifierKeys = 0;
                    _vkey = 0;
                    TextboxShotcut.Text = string.Empty;
                    WindowTitle.Title = "YasoCut - 快捷键无效";
                    ButtonShotcut.IsEnabled = false;
                    return;
                }
                TextboxMs.IsEnabled = true;
                _isSettingShortCut = false;
                TextboxShotcut.IsEnabled = false;
                TextboxPrefix.IsEnabled = true;
                TextboxPath.IsEnabled = true;
                ButtonOpen.IsEnabled = true;
                ButtonSelect.IsEnabled = true;
                CheckBoxTitle.IsEnabled = true;
                ButtonShotcut.Content = "设置快捷键";
                _notifyIcon.Text = WindowTitle.Title = _checkComMsMenuItem.Checked ?
                    "YasoCut - 连续模式" :
                    _checkObModeMenuItem.Checked ?
                    "YasoCut - 爆发模式" :
                    "YasoCut";
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
                _notifyIcon.Text = WindowTitle.Title = "YasoCut - 设置快捷键";
                ButtonShotcut.Content = "保存快捷键";
                NativeMethods.UnregisterHotKey(_hwndHelper.Handle, _atom);
                TextboxMs.IsEnabled = false;
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
                Keyboard.IsKeyDown(Key.System))
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

            if (_modifierKeys != ModifierKeys.None)
            {
                TextboxShotcut.Text = $"{_modifierKeys}, {e.Key}";
            }
            else
            {
                TextboxShotcut.Text = $"{e.Key}";
            }


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

                if (!(yasocut.GetValue("ConMs") is int comMs))
                {
                    comMs = 1;
                    yasocut.SetValue("ConMs", comMs, RegistryValueKind.DWord);
                }
                TextboxMs.Text = comMs.ToString();

                if (!(yasocut.GetValue("ObCount") is int obCount))
                {
                    obCount = 1;
                    yasocut.SetValue("ObCount", obCount, RegistryValueKind.DWord);
                }
                TextboxOB.Text = obCount.ToString();


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
                    if (_modifierKeys != ModifierKeys.None)
                    {
                        TextboxShotcut.Text = $"{_modifierKeys}, {KeyInterop.KeyFromVirtualKey(_vkey)}";
                    }
                    else
                    {
                        TextboxShotcut.Text = $"{KeyInterop.KeyFromVirtualKey(_vkey)}";
                    }
                    _notifyIcon.Text = WindowTitle.Title = _checkComMsMenuItem.Checked ?
                          "YasoCut - 连续模式" :
                          _checkObModeMenuItem.Checked ?
                          "YasoCut - 爆发模式" :
                          "YasoCut";
                }
                else
                {
                    MessageBox.Show("注册全局快捷键失败！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    _modifierKeys = 0;
                    _vkey = 0;
                    TextboxShotcut.Text = string.Empty;
                    _notifyIcon.Text = WindowTitle.Title = "YasoCut - 设置快捷键";
                    ButtonShotcut.Content = "保存快捷键";
                    NativeMethods.UnregisterHotKey(_hwndHelper.Handle, _atom);
                    TextboxMs.IsEnabled = false;
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
                MessageBox.Show("前缀中包含非法字符。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
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

        private void TextboxNumber_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            switch (e.Key)
            {
                case Key.Back:
                case Key.CapsLock:
                case Key.Escape:
                case Key.Left:
                case Key.Up:
                case Key.Right:
                case Key.Down:
                case Key.Delete:
                case Key.D0:
                case Key.D1:
                case Key.D2:
                case Key.D3:
                case Key.D4:
                case Key.D5:
                case Key.D6:
                case Key.D7:
                case Key.D8:
                case Key.D9:
                case Key.NumPad0:
                case Key.NumPad1:
                case Key.NumPad2:
                case Key.NumPad3:
                case Key.NumPad4:
                case Key.NumPad5:
                case Key.NumPad6:
                case Key.NumPad7:
                case Key.NumPad8:
                case Key.NumPad9:
                case Key.NumLock:
                    break;
                default:
                    e.Handled = true;
                    break;
            }
        }



        private void TextboxMs_TextChanged(object sender, TextChangedEventArgs e)
        {
            NumberBoxChangedCore(TextboxMs, "ConMs");
        }

        private void TextboxOB_TextChanged(object sender, TextChangedEventArgs e)
        {
            NumberBoxChangedCore(TextboxOB, "ObCount");
        }

        private bool _setTextbox = false;
        private void NumberBoxChangedCore(TextBox textbox, string key)
        {
            if (!_inited)
            {
                return;
            }
            if (_setTextbox)
            {
                return;
            }
            _setTextbox = true;
            string str = textbox.Text;
            if (string.IsNullOrWhiteSpace(str) || !int.TryParse(str, out int value) || value <= 0)
            {
                textbox.Text = string.Empty;
                value = 1;
            }
            using (RegistryKey soft = Registry.CurrentUser.OpenSubKey("SOFTWARE", true))
            {
                RegistryKey yasocut = soft.OpenSubKey("YasoCut", true) ?? soft.CreateSubKey("YasoCut", true);
                yasocut.SetValue(key, value, RegistryValueKind.DWord);
                yasocut.Dispose();
            }
            _setTextbox = false;
        }
    }
}
