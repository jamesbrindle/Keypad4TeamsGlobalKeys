using System;
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
    }
}
