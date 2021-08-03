using System.Threading;

namespace Keypad4Teams
{
    public class SafeThreading
    {
        public static void SafeSleep(int ms)
        {
            try
            {
                Thread.Sleep(ms);
            }
            catch { }
        }
    }
}
