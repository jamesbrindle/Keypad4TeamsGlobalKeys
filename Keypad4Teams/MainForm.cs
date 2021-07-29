using Keypad4Teams.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Linq;

namespace Keypad4Teams
{
    public partial class MainForm : Form
    {
        #region Dll Imports
        private delegate bool EnumWindowProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern int ShowWindow(IntPtr hWnd, uint Msg);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, UInt32 wParam, UInt32 lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int GetAsyncKeyState(Int32 i);

        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        #endregion

        private const int SW_RESTORE = 9;
        private const uint SW_RESTORE_HEX = 0x09;

        private const UInt32 WM_SYSCOMMAND = 0x0112;
        private const UInt32 SC_RESTORE = 0xF120;

        private GlobalKeyboardHook _globalKeyboardHook;
        private NotifyIcon _trayIcon;
        private List<ProcessAndHandle> _processAndHandlerList;

        public MainForm()
        {
            InitializeComponent();

            _globalKeyboardHook = new GlobalKeyboardHook(new Keys[] { Keys.NumPad0, Keys.D0, Keys.Alt });
            _globalKeyboardHook.KeyboardPressed += OnKeyPressed;

            ShowInTaskbar = false;
            Hide();
            WindowState = FormWindowState.Minimized;

            _trayIcon = new NotifyIcon()
            {
                Icon = Resources.favicon,
                ContextMenu = new ContextMenu(new MenuItem[] {
                new MenuItem("Exit", Exit)
            }),
                Visible = true
            };
        }

        private void FocusTeams()
        {
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    var teamsProcesses = Process.GetProcessesByName("Teams");
                    var validTeamsHandles = new List<IntPtr>();
                    _processAndHandlerList = new List<ProcessAndHandle>();

                    Parallel.ForEach(teamsProcesses, teamsProcess =>
                    {
                        string title = GetWindowTitle(teamsProcess.MainWindowHandle);
                        if (!string.IsNullOrEmpty(title) && title.Contains("| Microsoft Teams"))
                        {
                            if (!validTeamsHandles.Contains(teamsProcess.MainWindowHandle))
                            {
                                validTeamsHandles.Add(teamsProcess.MainWindowHandle);
                                _processAndHandlerList.Add(new ProcessAndHandle
                                {
                                    ValidProcess = teamsProcess,
                                    ValidHandle = teamsProcess.MainWindowHandle
                                });
                            }
                        }

                        var childWindows = new WindowHandleInfo(teamsProcess.MainWindowHandle).GetAllChildHandles();
                        if (childWindows != null)
                        {
                            Parallel.ForEach(childWindows, childWindow =>
                            {
                                string childTitle = GetWindowTitle(childWindow);
                                if (!string.IsNullOrEmpty(childTitle) && childTitle.Contains("| Microsoft Teams"))
                                {
                                    if (!validTeamsHandles.Contains(childWindow))
                                    {
                                        validTeamsHandles.Add(childWindow);
                                        _processAndHandlerList.Add(new ProcessAndHandle
                                        {
                                            ValidProcess = teamsProcess,
                                            ValidHandle = childWindow
                                        });
                                    }
                                }
                            });
                        }
                    });

                    break;
                }
                catch { }
            }

            foreach (var processAndHandle in _processAndHandlerList)
            {
                string title = GetWindowTitle(processAndHandle.ValidHandle);
                Console.Out.WriteLine(title);
            }

            if (_processAndHandlerList != null && _processAndHandlerList.Count > 0)
            {
                var selectedProcessAndHandle = _processAndHandlerList.OrderByDescending(p => p.ValidProcess.StartTime).FirstOrDefault();
                if (selectedProcessAndHandle != null)
                {
                    try
                    {
                        SetForegroundWindow(selectedProcessAndHandle.ValidHandle);
                    }
                    catch { }

                    try
                    {
                        ShowWindowAsync(selectedProcessAndHandle.ValidHandle, SW_RESTORE);
                    }
                    catch { }

                    try
                    {
                        ShowWindow(selectedProcessAndHandle.ValidHandle, SW_RESTORE_HEX);
                    }
                    catch { }

                    try
                    {
                        ShowWindowAsync(selectedProcessAndHandle.ValidProcess.Handle, SW_RESTORE);
                    }
                    catch { }

                    try
                    {
                        ShowWindow(selectedProcessAndHandle.ValidProcess.Handle, SW_RESTORE_HEX);
                    }
                    catch { }

                    try
                    {
                        SendMessage(selectedProcessAndHandle.ValidHandle, WM_SYSCOMMAND, SC_RESTORE, 0);
                    }
                    catch { }

                    try
                    {
                        SetForegroundWindow(selectedProcessAndHandle.ValidHandle);
                    }
                    catch { }
                }
            }
        }

        public static string GetWindowTitle(IntPtr hWnd)
        {
            var length = GetWindowTextLength(hWnd) + 1;
            var title = new StringBuilder(length);
            GetWindowText(hWnd, title, length);
            return title.ToString();
        }

        private void OnKeyPressed(object sender, GlobalKeyboardHookEventArgs e)
        {
            if (e.KeyboardState == GlobalKeyboardHook.KeyboardState.SysKeyDown &&
                 (e.KeyboardData.Key == Keys.D0 || e.KeyboardData.Key == Keys.NumPad0 ||
                  e.KeyboardData.VirtualCode == 96 || e.KeyboardData.VirtualCode == 48))
            {
                FocusTeams();
            }
        }

        private void Exit(object sender, EventArgs e)
        {
            _trayIcon.Visible = false;
            Application.Exit();
        }

        public new void Dispose()
        {
            _globalKeyboardHook?.Dispose();
        }
    }
}
