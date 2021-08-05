using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;
using Keypad4Teams.Properties;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using static Keypad4Teams.WindowHelper;

namespace Keypad4Teams
{
    public partial class MainForm : Form
    {   
        private NotifyIcon _trayIcon;
        private IntPtr _altHandler = IntPtr.Zero;

        private readonly int _msgNotify;
        public delegate void EventHandler(object sender, int action, IntPtr handle);

        public event EventHandler WindowEvent;
        private GlobalKeyboardHook GlobalKeyboardHook { get; set; }
        private List<ProcessAndHandle> RecentProcessAndHandle { get; set; }
        public List<ProcessAndHandle> HandleCache { get; set; } = new List<ProcessAndHandle>();
        private System.Windows.Forms.Timer CheckCacheTimer { get; set; } = new System.Windows.Forms.Timer();

        public MainForm()
        {
            InitializeComponent();

            GlobalKeyboardHook = new GlobalKeyboardHook(new Keys[] { Keys.NumPad0, Keys.D0, Keys.Alt });
            GlobalKeyboardHook.KeyboardPressed += OnKeyPressed;

            SetStartWithWindows();

            ShowInTaskbar = false;
            FormBorderStyle = FormBorderStyle.None;
            WindowState = FormWindowState.Minimized;
            Opacity = 0;
            Hide();

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

            _msgNotify = RegisterWindowMessage("SHELLHOOK");
            RegisterShellHookWindow(Handle);

            WindowEvent += WindowEventHandler;

            CheckCacheTimer.Interval = 20000; // 20 seconds
            CheckCacheTimer.Tick += new System.EventHandler(CheckCacheTimer_Tick);
            CheckCacheTimer.Start();
        }

        public void WindowEventHandler(object sender, int action, IntPtr handle)
        {
            switch ((ShellEvents)action)
            {
                case ShellEvents.HSHELL_WINDOWCREATED:
                    new Thread((ThreadStart)delegate
                    {
                        SafeThreading.SafeSleep(500);
                        string windowTitle = GetWindowTitle(handle);

                        int it = 0;
                        while (windowTitle == "Teams" && it < 50)
                        {
                            SafeThreading.SafeSleep(100);
                            windowTitle = GetWindowTitle(handle);
                            it++;
                        }

                        try
                        {
                            if (!HandleCache.Any(m => m.Handle == handle))
                            {                               
                                if (windowTitle != null && windowTitle.Length > 0 && !windowTitle.ToLower().Contains("call in progress"))
                                {
                                    HandleCache.Add(new ProcessAndHandle
                                    {
                                        Handle = handle,
                                        WindowTitle = windowTitle,
                                        IsCallWindow = IsCallWindow(windowTitle, handle, out var process, out bool nullHandle, out var elements),
                                        NullHandle = nullHandle,
                                        Elements = elements,
                                        ValidProcess = process
                                    });
                                }
                            }
                        }
                        catch { }
                    }).Start();
                    break;
                case ShellEvents.HSHELL_WINDOWDESTROYED:
                    try
                    {
                        HandleCache.RemoveAll(m => m.Handle == handle);
                    }
                    catch { }
                    break;
            }
        }

        protected virtual void OnWindowEvent(int action, IntPtr handle)
        {
            var handler = WindowEvent;
            if (handler != null)
                handler(this, action, handle);
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
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == _msgNotify)
            {
                switch ((ShellEvents)m.WParam.ToInt32())
                {
                    case ShellEvents.HSHELL_WINDOWCREATED:
                    case ShellEvents.HSHELL_WINDOWDESTROYED:
                    case ShellEvents.HSHELL_WINDOWACTIVATED:
                        string processTitle = GetWindowTitle(m.LParam);
                        if (processTitle.ToLower().Contains("teams"))
                        {
                            int action = m.WParam.ToInt32();
                            OnWindowEvent(action, m.LParam);
                        }
                        break;
                }
            }
            base.WndProc(ref m);
        }

        private void CheckCacheTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                var removeList = new List<ProcessAndHandle>();
                foreach (var ph in HandleCache)
                {
                    if (!IsWindow(ph.Handle))
                        removeList.Add(ph);
                }

                foreach (var ph in removeList)
                {
                    try
                    {
                        HandleCache.RemoveAll(m => m.Handle == ph.Handle);
                    }
                    catch { }
                }
            }
            catch { }
        }

        private void FocusTeams()
        {
            try
            {
                bool isCachedValid = false;
                if (HandleCache.Count(m => m.IsCallWindow) == 1)
                {
                    var cph = HandleCache.Where(m => m.IsCallWindow).FirstOrDefault();
                    if (IsWindow(cph.Handle))
                    {
                        RecentProcessAndHandle = new List<ProcessAndHandle>
                        {
                            cph
                        };

                        isCachedValid = true;
                    }
                }

                if (!isCachedValid)
                    GatherPotentialTeamsWindows();

                if (RecentProcessAndHandle != null && RecentProcessAndHandle.Count > 0)
                {
                    ProcessAndHandle selectedProcessAndHandle = null;
                    if (RecentProcessAndHandle.Count(m => m.IsCallWindow) > 1)
                        selectedProcessAndHandle = SelectHandleWhenTwoWindowsReportIsCall();

                    else if (RecentProcessAndHandle.Count(m => !m.IsCallWindow) == RecentProcessAndHandle.Count)
                        selectedProcessAndHandle = SelectHandleWhenNoWindowsReportIsCall();

                    else
                        selectedProcessAndHandle = RecentProcessAndHandle.OrderByDescending(p => p.IsCallWindow)
                                                                         .ThenBy(p => p.NullHandle)
                                                                         .FirstOrDefault();
                    if (selectedProcessAndHandle != null)
                        AggressiveSetForgroundWindow(selectedProcessAndHandle.Handle);
                }
            }
            catch { }

            _altHandler = IntPtr.Zero;
        }

        private void GatherPotentialTeamsWindows()
        {
            for (int i = 0; i < 2; i++)
            {
                try
                {
                    var teamsProcesses = Process.GetProcessesByName("Teams");
                    var validTeamsHandles = new List<IntPtr>();
                    RecentProcessAndHandle = new List<ProcessAndHandle>();

                    foreach (var teamsProcess in teamsProcesses)
                    {
                        string title = GetWindowTitle(teamsProcess.MainWindowHandle);
                        if (!string.IsNullOrEmpty(title) && title.Contains("| Microsoft Teams"))
                        {
                            if (!validTeamsHandles.Contains(teamsProcess.MainWindowHandle))
                            {
                                validTeamsHandles.Add(teamsProcess.MainWindowHandle);
                                var ph = new ProcessAndHandle
                                {
                                    ValidProcess = teamsProcess,
                                    Handle = teamsProcess.MainWindowHandle,
                                    WindowTitle = title,
                                    IsCallWindow = IsCallWindow(title, teamsProcess.MainWindowHandle, out var process, out bool nullHandle, out var elements),
                                    Elements = elements,
                                    NullHandle = nullHandle
                                };

                                RecentProcessAndHandle.Add(ph);
                                if (!HandleCache.Any(m => m.Handle == teamsProcess.MainWindowHandle))
                                    HandleCache.Add(ph);
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
                                        var ph = new ProcessAndHandle
                                        {
                                            ValidProcess = teamsProcess,
                                            Handle = childWindow,
                                            WindowTitle = title,
                                            IsCallWindow = IsCallWindow(title, childWindow, out var process, out bool nullHandle, out var elements),
                                            Elements = elements,
                                            NullHandle = nullHandle
                                        };

                                        RecentProcessAndHandle.Add(ph);
                                        if (!HandleCache.Any(m => m.Handle == teamsProcess.MainWindowHandle))
                                            HandleCache.Add(ph);
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

        private bool IsCallWindow(
            string windowName, 
            IntPtr handle, 
            out Process process, 
            out bool nullHandle, 
            out List<AutomationElement> elements)
        {
            nullHandle = false;
            elements = new List<AutomationElement>();
            process = null;

            if (HandleCache.Any(m => m.Handle == handle))
            {
                var ph = HandleCache.Where(m => m.Handle == handle).FirstOrDefault();
                elements = ph.Elements;
                process = ph.ValidProcess;
                nullHandle = ph.NullHandle;

                return ph.IsCallWindow;
            }
            else
            {
                try
                {
                    using (var automation = new UIA3Automation())
                    {
                        try
                        {
                            var desktop = automation.GetDesktop();
                            var children = desktop.FindAllChildren(c => c.ByName(windowName));
                            var child = children.Where(m => m.Properties.NativeWindowHandle == handle).FirstOrDefault();

                            if (child == null)
                            {
                                nullHandle = true;
                                return false;
                            }

                            process = Process.GetProcessById(child.Properties.ProcessId);
                            if (_altHandler == IntPtr.Zero)
                            {
                                try
                                {
                                    if (child != null && child.AutomationId == null)
                                        _altHandler = child.Properties.NativeWindowHandle;
                                }
                                catch
                                {
                                    try
                                    {
                                        if (child != null && child.Properties != null && child.Properties.NativeWindowHandle != null)
                                            _altHandler = child.Properties.NativeWindowHandle;
                                    }
                                    catch { }
                                }
                            }

                            List<AutomationElement> elementsList = null;
                            GetAllElementsRecurisve(child, ref elementsList);
                            elements = elementsList;

                            for (int i = 0; i < elements.Count; i++)
                            {
                                try
                                {
                                    if (elements[i].AutomationId == "prejoin-devicesettings-button" ||
                                        elements[i].AutomationId == "prejoin-devicesettings-button" ||
                                        elements[i].AutomationId == "prejoin-join-button" ||
                                        elements[i].Name == "Microphone" ||
                                        elements[i].Name == "ComputerAudio" ||
                                        elements[i].Name == "Open device settings" ||
                                        elements[i].AutomationId == "hangup-button" ||
                                        elements[i].Name == "Leave" ||
                                        elements[i].Name == "video options" ||
                                        elements[i].Name == "Leave" ||
                                        elements[i].AutomationId == "microphone-button" ||
                                        elements[i].Name == "Mute" ||
                                        elements[i].AutomationId == "video-button" ||
                                        elements[i].Name == "Turn Camera On" ||
                                        elements[i].Name == "Turn Camera Off")
                                    {
                                        return true;
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }

            return false;
        }

        private ProcessAndHandle SelectHandleWhenTwoWindowsReportIsCall()
        {
            return new ProcessAndHandle
            {
                ValidProcess = RecentProcessAndHandle.FirstOrDefault().ValidProcess,
                WindowTitle = RecentProcessAndHandle.FirstOrDefault().WindowTitle,
                Handle = _altHandler
            };
        }

        private ProcessAndHandle SelectHandleWhenNoWindowsReportIsCall()
        {
            // Use a points system to try and 'guess' the call Window

            foreach (var ph in RecentProcessAndHandle)
            {
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
            foreach (var ph in RecentProcessAndHandle)
            {
                if (ph.Points > max)
                    max = ph.Points;
            }

            int maxCount = 0;
            foreach (var ph in RecentProcessAndHandle)
            {
                if (ph.Points == max)
                    maxCount++;
            }

            if (maxCount == 1)
            {
                return RecentProcessAndHandle.OrderByDescending(m => m.Points)
                                             .ThenBy(m => m.NullHandle)
                                             .FirstOrDefault();
            }
            else
            {
                var rand = new Random();
                var possibles = RecentProcessAndHandle.Where(m => m.Points == maxCount).ToList();
                int toSkip = rand.Next(0, possibles.Count);

                return possibles.Skip(toSkip).Take(1).First();
            }
        }

        private void GetAllElementsRecurisve(
            AutomationElement parent,
            ref List<AutomationElement> elements,
            int count = 0)
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
                        if (element.AutomationId == "prejoin-devicesettings-button" ||
                            element.AutomationId == "prejoin-devicesettings-button" ||
                            element.AutomationId == "prejoin-join-button" ||
                            element.Name == "Microphone" ||
                            element.Name == "ComputerAudio" ||
                            element.Name == "Open device settings" ||
                            element.AutomationId == "hangup-button" ||
                            element.Name == "Leave" ||
                            element.Name == "video options" ||
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
                var cp = base.CreateParams;
                cp.ExStyle |= 0x80;
                return cp;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Dispose();
        }

        public new void Dispose()
        {
            try
            {
                CheckCacheTimer.Stop();
            }
            catch { }

            try
            {
                DeregisterShellHookWindow(Handle);
            }
            catch { }

            try
            {
                GlobalKeyboardHook?.Dispose();
            }
            catch { }
        }
    }
}
