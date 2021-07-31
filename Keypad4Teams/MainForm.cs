﻿using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Keypad4Teams.Properties;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

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
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int GetAsyncKeyState(Int32 i);

        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        private struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }

        #endregion

        private const int SW_RESTORE = 9;
        private const uint SW_RESTORE_HEX = 0x09;

        private const UInt32 WM_SYSCOMMAND = 0x0112;
        private const UInt32 SC_RESTORE = 0xF120;

        private GlobalKeyboardHook _globalKeyboardHook;
        private NotifyIcon _trayIcon;
        private List<ProcessAndHandle> _processAndHandlerList;
        private IntPtr _altHandler = IntPtr.Zero;
        private int _iterationIndex = 0;

        public MainForm()
        {
            InitializeComponent();

            _globalKeyboardHook = new GlobalKeyboardHook(new Keys[] { Keys.NumPad0, Keys.D0, Keys.Alt });
            _globalKeyboardHook.KeyboardPressed += OnKeyPressed;

            SetStartWithWindows();

            ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.None;
            Opacity = 0;
            Hide();

            Hide();
            WindowState = FormWindowState.Minimized;

            _trayIcon = new NotifyIcon()
            {
                Icon = Resources.favicon,
                Text = "Keypad4Teams",
                ContextMenu = new ContextMenu(new System.Windows.Forms.MenuItem[]
                {
                    new System.Windows.Forms.MenuItem("Exit", Exit),
                }),
                Visible = true
            };
        }

        private void FocusTeams()
        {
            try
            {
                GatherPotentialTeamsWindows();

                if (_processAndHandlerList != null && _processAndHandlerList.Count > 0)
                {
                    ProcessAndHandle selectedProcessAndHandle = null;

                    if (_processAndHandlerList.Count(m => m.IsCallWindow) > 1)
                        selectedProcessAndHandle = SelectHandleWhenTwoWindowsReportIsCall();

                    else if (_processAndHandlerList.Count(m => !m.IsCallWindow) == _processAndHandlerList.Count)
                        selectedProcessAndHandle = SelectHandleWhenNoWindowsReportIsCall();

                    else
                        selectedProcessAndHandle = _processAndHandlerList.OrderByDescending(p => p.IsCallWindow)
                                                                         .ThenBy(p => p.NullHandle)
                                                                         .FirstOrDefault();
                    if (selectedProcessAndHandle != null)
                    {
                        try
                        {
                            SetForegroundWindow(selectedProcessAndHandle.ValidHandle);
                        }
                        catch { }

                        try
                        {
                            if (IsWindowMinimised(selectedProcessAndHandle.ValidHandle))
                                ShowWindowAsync(selectedProcessAndHandle.ValidHandle, SW_RESTORE);
                        }
                        catch { }

                        try
                        {
                            if (IsWindowMinimised(selectedProcessAndHandle.ValidHandle))
                                ShowWindow(selectedProcessAndHandle.ValidHandle, SW_RESTORE_HEX);
                        }
                        catch { }

                        try
                        {
                            if (IsWindowMinimised(selectedProcessAndHandle.ValidHandle))
                                ShowWindowAsync(selectedProcessAndHandle.ValidProcess.Handle, SW_RESTORE);
                        }
                        catch { }

                        try
                        {
                            if (IsWindowMinimised(selectedProcessAndHandle.ValidHandle))
                                ShowWindow(selectedProcessAndHandle.ValidProcess.Handle, SW_RESTORE_HEX);
                        }
                        catch { }

                        try
                        {
                            if (IsWindowMinimised(selectedProcessAndHandle.ValidHandle))
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
            catch { }

            _altHandler = IntPtr.Zero;
            _iterationIndex = 0;
        }

        private void GatherPotentialTeamsWindows()
        {
            for (int i = 0; i < 2; i++)
            {
                try
                {
                    var teamsProcesses = Process.GetProcessesByName("Teams");
                    var validTeamsHandles = new List<IntPtr>();
                    _processAndHandlerList = new List<ProcessAndHandle>();

                    foreach (var teamsProcess in teamsProcesses)
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
                                    ValidHandle = teamsProcess.MainWindowHandle,
                                    WindowTitle = title,
                                    IsCallWindow = IsCallWindow(title, out bool nullHandle, out List<AutomationElement> elements, out bool blackPixel),
                                    BlackPixel = blackPixel,
                                    Elements = elements,
                                    NullHandle = nullHandle,
                                    IterationIndex = _iterationIndex
                                });
                                _iterationIndex++;
                            }
                        }

                        var childWindows = new ChildWindowHandler(teamsProcess.MainWindowHandle).GetAllChildHandles();
                        if (childWindows != null)
                        {
                            foreach (var childWindow in childWindows)
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
                                            ValidHandle = childWindow,
                                            WindowTitle = childTitle,
                                            IsCallWindow = IsCallWindow(childTitle, out bool nullHandle, out List<AutomationElement> elements, out bool blackPixel),
                                            BlackPixel = blackPixel,
                                            Elements = elements,
                                            NullHandle = nullHandle,
                                            IterationIndex = _iterationIndex
                                        });
                                        _iterationIndex++;
                                    }
                                }
                            }
                        }
                    }

                    break;
                }
                catch { }
            }
        }

        private bool IsWindowMinimised(IntPtr handle)
        {
            if (handle != IntPtr.Zero)
            {
                WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
                GetWindowPlacement(handle, ref placement);
                switch (placement.showCmd)
                {
                    case 2:
                        return true;
                }
            }

            return false;
        }

        private bool IsCallWindow(string windowName, out bool nullHandle, out List<AutomationElement> elements, out bool blackPixel)
        {
            nullHandle = false;
            elements = new List<AutomationElement>();
            blackPixel = false;

            try
            {
                using (var automation = new UIA3Automation())
                {
                    try
                    {
                        var desktop = automation.GetDesktop();
                        var parent = desktop.FindFirstChild(c => c.ByName(windowName));

                        if (parent == null)
                        {
                            nullHandle = true;
                            return false;
                        }

                        if (_altHandler == IntPtr.Zero)
                        {
                            try
                            {
                                if (parent != null && parent.AutomationId == null)
                                    _altHandler = parent.Properties.NativeWindowHandle;
                            }
                            catch
                            {
                                try
                                {
                                    if (parent != null && parent.Properties != null && parent.Properties.NativeWindowHandle != null)
                                        _altHandler = parent.Properties.NativeWindowHandle;
                                }
                                catch { }
                            }
                        }

                        try
                        {
                            var bmp = parent.Capture();
                            var pixelColor = bmp.GetPixel(1, 1);

                            if (pixelColor.R <= 22 && pixelColor.G <= 22 && pixelColor.B <= 22)
                                blackPixel = true;

                            bmp.Dispose();
                        }
                        catch { }

                        List<AutomationElement> elementsList = null;
                        GetAllElementsRecurisve(parent, ref elementsList);
                        elements = elementsList;

                        for (int i = 0; i < elements.Count; i++)
                        {
                            try
                            {
                                if (elements[i].AutomationId == "hangup-button" ||
                                    elements[i].Name == "Leave" ||
                                    elements[i].AutomationId == "microphone-button" ||
                                    elements[i].Name == "Mute" ||
                                    elements[i].AutomationId == "video-button" ||
                                    elements[i].Name == "Turn Camera On" ||
                                    elements[i].Name == "Turn Camera Off")
                                    return true;
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }

            return false;
        }

        private ProcessAndHandle SelectHandleWhenTwoWindowsReportIsCall()
        {
            return new ProcessAndHandle
            {
                ValidProcess = _processAndHandlerList.FirstOrDefault().ValidProcess,
                WindowTitle = _processAndHandlerList.FirstOrDefault().WindowTitle,
                ValidHandle = _altHandler
            };
        }

        private ProcessAndHandle SelectHandleWhenNoWindowsReportIsCall()
        {
            // Use a points system to try and 'guess' the call Window

            foreach (var ph in _processAndHandlerList)
            {
                if (ph.BlackPixel)
                    ph.Points++;

                if (ph.Elements.Count >= 2)
                {
                    int hasValCount = 0;
                    for (int i = 0; i < ph.Elements.Count; i++)
                    {
                        try
                        {
                            if (ph.Elements[i].AutomationId != null && ph.Elements[i].AutomationId.Length > 0)
                                hasValCount++;
                        }
                        catch { }
                    }

                    if (hasValCount < 2)
                        ph.Points++;
                }

                if (ph.Elements.Count == 2 || (ph.Elements.Count >= 6 && ph.Elements.Count <= 7))
                    ph.Points++;

                if (ph.Elements.Count > 30)
                    ph.Points--;

                if (ph.Elements.Count == 0)
                    ph.Points--;
            }

            int max = -99;

            foreach (var ph in _processAndHandlerList)
            {
                if (ph.Points > max)
                    max = ph.Points;
            }

            int maxCount = 0;
            foreach (var ph in _processAndHandlerList)
            {
                if (ph.Points == max)
                    maxCount++;
            }

            if (maxCount == 1)
            {
                return _processAndHandlerList.OrderByDescending(m => m.Points)
                                             .ThenBy(m => m.NullHandle)
                                             .FirstOrDefault();
            }
            else
            {
                var rand = new Random();
                var possibles = _processAndHandlerList.Where(m => m.Points == maxCount).ToList();
                int toSkip = rand.Next(0, possibles.Count);

                return possibles.Skip(toSkip).Take(1).First();
            }
        }

        private void GetAllElementsRecurisve(AutomationElement parent, ref List<AutomationElement> elements, int count = 0)
        {
            if (elements == null)
                elements = new List<AutomationElement>();

            if (count > 50)
                return;

            if (parent != null)
            {
                foreach (var element in parent.FindAllChildren())
                {
                    elements.Add(element);
                    count++;

                    if (count > 50)
                        break;

                    try
                    {
                        if (element.AutomationId == "hangup-button" ||
                            element.Name == "Leave" ||
                            element.AutomationId == "microphone-button" ||
                            element.Name == "Mute" ||
                            element.AutomationId == "video-button" ||
                            element.Name == "Turn Camera On" ||
                            element.Name == "Turn Camera Off")
                            break;
                    }
                    catch { }

                    GetAllElementsRecurisve(element, ref elements, count);
                }

                if (count > 50)
                    return;
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

        public void SetStartWithWindows()
        {
            try
            {
                string path = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
                var key = Registry.CurrentUser.OpenSubKey(path, true);
                key.SetValue("Keypad4Teams", "\"" + Application.ExecutablePath.ToString() + "\"");
            }
            catch { }
        }

        private void Exit(object sender, EventArgs e)
        {
            _trayIcon.Visible = false;
            Application.Exit();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                // turn on WS_EX_TOOLWINDOW style bit
                cp.ExStyle |= 0x80;
                return cp;
            }
        }

        public new void Dispose()
        {
            _globalKeyboardHook?.Dispose();
        }
    }
}
