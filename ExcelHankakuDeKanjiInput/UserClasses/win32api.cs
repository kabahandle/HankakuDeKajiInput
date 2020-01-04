using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace MyVSNetAddin2.UserClasses
{
    public class win32api
    {

        [DllImport("user32.dll", EntryPoint = "GetCaretPos")]
        static extern bool GetCaretPos(out Point lpPoint);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern IntPtr AttachThreadInput(IntPtr idAttach, IntPtr idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, IntPtr lpdwProcessId);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetCurrentThreadId();

        [DllImport("user32.dll")]
        static extern IntPtr GetFocus();

        [DllImport("user32.dll")]
        static extern bool ClientToScreen(IntPtr hwnd, out Point lpPoint);

        public static void Pos(ref Point p)
        {
            IntPtr hWnd = GetForegroundWindow();

            IntPtr current = (IntPtr)GetCurrentThreadId();
            IntPtr target = (IntPtr)GetWindowThreadProcessId(hWnd, IntPtr.Zero);


            AttachThreadInput(current, target, true);
            GetCaretPos(out p);
            IntPtr fWnd = GetFocus();
            ClientToScreen(fWnd, out p);
           // ClientToScreen(hWnd, out p);
            AttachThreadInput(current, target, false);
        }
    }
}
