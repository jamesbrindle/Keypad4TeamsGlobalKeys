using FlaUI.Core.AutomationElements;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Keypad4Teams
{
    public class ProcessAndHandle
    {
        public Process ValidProcess { get; set; }
        public IntPtr ValidHandle { get; set; }
        public string WindowTitle { get; set; }
        public bool IsCallWindow { get; set; } = false;
        public bool NullHandle { get; set; } = false;
        public int IterationIndex { get; set; } = 0;
        public List<AutomationElement> Elements { get; set; } = new List<AutomationElement>();
        public int Points { get; set; } = 0;
        public bool BlackPixel { get; set; } = false;
    }
}
