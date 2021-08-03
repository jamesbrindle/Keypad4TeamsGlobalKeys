﻿using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace Keypad4Teams
{
    public class WindowHelper
    {
        #region Dll Imports

        private delegate bool EnumWindowProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern int ShowWindow(IntPtr hWnd, uint Msg);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, UInt32 wParam, UInt32 lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern int GetAsyncKeyState(Int32 i);

        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [DllImport("user32.dll", EntryPoint = "RegisterWindowMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int RegisterWindowMessage(string lpString);
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]

        public static extern int DeregisterShellHookWindow(IntPtr hWnd);
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern int RegisterShellHookWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        #endregion

        public struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }

        public enum ShellEvents : int
        {
            HSHELL_WINDOWCREATED = 1,
            HSHELL_WINDOWDESTROYED = 2,
            HSHELL_ACTIVATESHELLWINDOW = 3,
            HSHELL_WINDOWACTIVATED = 4,
            HSHELL_GETMINRECT = 5,
            HSHELL_REDRAW = 6,
            HSHELL_TASKMAN = 7,
            HSHELL_LANGUAGE = 8,
            HSHELL_ACCESSIBILITYSTATE = 11,
            HSHELL_APPCOMMAND = 12
        }

        public const int SW_RESTORE = 9;
        public const uint SW_RESTORE_HEX = 0x09;

        public const uint WM_SYSCOMMAND = 0x0112;
        public const uint SC_RESTORE = 0xF120;

        public static string GetWindowTitle(IntPtr hWnd)
        {
            int length = GetWindowTextLength(hWnd) + 1;
            var title = new StringBuilder(length);
            GetWindowText(hWnd, title, length);
            return title.ToString();
        }

        public static bool IsWindowMinimised(IntPtr handle)
        {
            if (handle != IntPtr.Zero)
            {
                var placement = new WINDOWPLACEMENT();
                GetWindowPlacement(handle, ref placement);
                switch (placement.showCmd)
                {
                    case 2:
                        return true;
                }
            }

            return false;
        }

        public class ChildWindowHandler
        {
            private delegate bool EnumWindowProc(IntPtr hwnd, IntPtr lParam);

            [DllImport("user32")]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr lParam);

            private IntPtr _MainHandle;

            public ChildWindowHandler(IntPtr handle)
            {
                this._MainHandle = handle;
            }

            public List<IntPtr> GetAllChildHandles()
            {
                var childHandles = new List<IntPtr>();

                var gcChildhandlesList = GCHandle.Alloc(childHandles);
                var pointerChildHandlesList = GCHandle.ToIntPtr(gcChildhandlesList);

                try
                {
                    var childProc = new EnumWindowProc(EnumWindow);
                    EnumChildWindows(this._MainHandle, childProc, pointerChildHandlesList);
                }
                finally
                {
                    gcChildhandlesList.Free();
                }

                return childHandles;
            }

            private bool EnumWindow(IntPtr hWnd, IntPtr lParam)
            {
                var gcChildhandlesList = GCHandle.FromIntPtr(lParam);

                if (gcChildhandlesList == null || gcChildhandlesList.Target == null)
                {
                    return false;
                }

                var childHandles = gcChildhandlesList.Target as List<IntPtr>;
                childHandles.Add(hWnd);

                return true;
            }
           
        }
    }
}
