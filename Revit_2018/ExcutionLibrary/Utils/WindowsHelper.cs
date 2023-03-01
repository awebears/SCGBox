using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Revit_2018.ExcutionLibrary.Utils
{
    internal static class WindowsHelper
    {

        //暂时不懂，网上抄的
        [DllImport("user32.dll", CharSet = CharSet.None, ExactSpelling = false)]
        public static extern bool EnumChildWindows(IntPtr hwndParent, CallBack lpEnumFunc, IntPtr lParam);
        public delegate bool CallBack(IntPtr hwnd, int lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpText, int nCount);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll", EntryPoint = "SendMessageA")]
        public static extern int SendMessage(IntPtr hwnd, uint wMsg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern void Keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtralInfo);
        [DllImport("user32.dll")]
        public static extern byte MapVirtualKey(byte wCode, int wMap);
    }
}