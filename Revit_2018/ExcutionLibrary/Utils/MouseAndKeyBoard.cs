using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using Gma.System.MouseKeyHook;
using System.Windows.Forms;

namespace Revit_2018.ExcutionLibrary.Utils
{
    internal class MouseAndKeyBoard
    {
        //private IKeyboardMouseEvents MouseEvents;
        private IKeyboardMouseEvents KeyBoardEvents;

        public void MouseAndKeyBoard_Subscribe()
        {
            //MouseEvents = Hook.GlobalEvents();
            //MouseEvents.MouseClick += GlobalHookMouseClick_Right;//鼠标点击事件
            //MouseEvents.MouseDoubleClick += GlobalHookMouseDoubleClick;//鼠标双击事件----不适用于revit，会导致当前选择的丢失
            //MouseEvents.MouseDownExt += GlobalHookMouseDownExt;

            KeyBoardEvents = Hook.GlobalEvents();
            KeyBoardEvents.KeyPress += SpaceKeyPress;//空格事件

        }

        private void SpaceKeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Space) { SendComplete(); };
        }

        public void MouseAndKeyBoard_Unsubscribe()
        {
            //MouseEvents.MouseClick -= GlobalHookMouseClick_Right;
            //MouseEvents.MouseDoubleClick -= GlobalHookMouseDoubleClick;
            //MouseEvents.MouseDownExt -= GlobalHookMouseDownExt;

            KeyBoardEvents.KeyPress -= SpaceKeyPress;//空格事件
            KeyBoardEvents.Dispose();
        }

        private void GlobalHookMouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left) { SendComplete(); };
        }

        private void GlobalHookMouseClick_Right(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right) { SendComplete(); };
        }


        private void SendComplete()
        {
            IntPtr revitWindow = Autodesk.Windows.ComponentManager.ApplicationWindow;//在AdWindows包里
            List<IntPtr> intPtrs = new List<IntPtr>();
            bool flag = WindowsHelper.EnumChildWindows(revitWindow,
                (hwnd, l) =>
                {
                    StringBuilder windowText = new StringBuilder(200);
                    WindowsHelper.GetWindowText(hwnd, windowText, windowText.Capacity);
                    StringBuilder className = new StringBuilder(200);
                    WindowsHelper.GetClassName(hwnd, className, className.Capacity);
                    if ((windowText.ToString().Equals("完成", StringComparison.Ordinal) ||
                         windowText.ToString().Equals("Finish", StringComparison.Ordinal)) &&
                         className.ToString().Contains("Button"))
                    {
                        intPtrs.Add(hwnd);
                        return false;
                    }
                    return true;
                }
                , new IntPtr(0));
            IntPtr complete = intPtrs.FirstOrDefault();
            WindowsHelper.SendMessage(complete, 245, 0, 0);
        }
    }
}