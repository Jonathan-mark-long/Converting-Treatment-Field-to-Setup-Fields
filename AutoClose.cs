using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CloseErrors
{
    public class WindowClose
    {
        public static string GetControlText(IntPtr textpointer)
        {
            StringBuilder sb = new StringBuilder(256);
            GetWindowCaption(textpointer, sb, 256);
            return sb.ToString();
        }
        public void ClosingWindow()
        {
            IntPtr messageBoxPointer = new IntPtr();
            IntPtr desktopPtr = GetDesktopWindow();

            for (int i = 0; i < 2000; i++)
            {

                messageBoxPointer = FindWindowsEx(desktopPtr, messageBoxPointer, "#32770", null);
                IntPtr textPointer = new IntPtr();

                for (int j = 0; j < 2000; j++)
                {
                    textPointer = FindWindowsEx(messageBoxPointer, textPointer, "Static", null);

                    if (textPointer == new IntPtr())
                    {
                        if (!i.Equals(0)) break;
                    }
                    else
                    {
                        string foundText = GetControlText(textPointer);
                        if (foundText.Contains("The following parameters were adjusted") || foundText.Contains("The field exits the patient support device in a region") || foundText.Contains("The field enters the patient support device in a region") || foundText.Contains("CUDA Run time error: invalid argument"))
                        {
                            const int WM_CLOSE = 0x0010;
                            const int TIMEOUT_INT = 2000;
                            SendMessageTimeout(messageBoxPointer, WM_CLOSE, IntPtr.Zero, IntPtr.Zero, SendMessageTimeoutFlags.SMTO_ABORTIFHUNG, TIMEOUT_INT, out UIntPtr test);
                        }
                    }
                }

            }

        }


        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        public static extern IntPtr FindWindowsEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpClassName, string lpszWindow);

        [DllImport("user32.dll", SetLastError = false)]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll", EntryPoint = "GetWindowText", CharSet = CharSet.Auto)]
        public static extern IntPtr GetWindowCaption(IntPtr hwnd, StringBuilder lpString, int maxCount);

        public enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0,
            SMTO_BLOCK = 0x1,
            SMTO_ABORTIFHUNG = 0x2,
            SMTO_NOTIMEOUTIFNOTHUNG = 0x8,
            SMTO_ERRORONEXIT = 0x20
        }

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, SendMessageTimeoutFlags Flags, uint uTimeout, out UIntPtr lpdwResult);



    }
}
