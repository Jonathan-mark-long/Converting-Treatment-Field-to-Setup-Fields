using ExecutableLogic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using CloseErrors;
using System.Threading;
using System.Windows.Forms;
using Application = VMS.TPS.Common.Model.API.Application;
using System.Runtime.InteropServices;

[assembly: ESAPIScript(IsWriteable = true)]

namespace MoveMouse
{
    class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            try
            {
                using (Application app = Application.CreateApplication())
                {


                    Perform(app, args);


                }

            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
                Console.Read();
            }
        }
        public static void Perform(Application app, string[] args)
        {

            ScriptContextArgs ctx = ScriptContextArgs.From(args);

            Patient patient = app.OpenPatientById(ctx.PatientId);
            string patientID = patient.Id;
            Course course = patient.Courses.First(e => e.Id == ctx.CourseId);
            ExternalPlanSetup plan = course.ExternalPlanSetups.First(e => e.Id == ctx.PlanSetupId);
            List<Beam> beams = plan.Beams.Where(b => b.Id.ToLower().Contains("setup")).OrderBy(b => b.BeamNumber).ToList();

            Tuple<int, int, int, int> window = WindowTuple(patientID);
            int left = window.Item1;
            int width = window.Item2;
            int top = window.Item3;
            int height = window.Item4;


            int LRScroll = 0;
            int TBSCroll = 0;
            int LRItem = 0;
            int TBItem = 0;

            APPBARDATA data = new APPBARDATA();
            data.cbSize = Marshal.SizeOf(data);
            SHAppBarMessage(ABM_GETTASKBARPOS, ref data);

            // Get the size of the taskbar icons from the RECT structure
            int iconSize = data.rc.Bottom - data.rc.Top;

            if (width < 1940)
            {
                if (iconSize == 40)
                {

                    LRScroll = (int)(left + (width - width * 0.845));
                    TBSCroll = (int)(top + (height - height * 0.245));
                    LRItem = (int)(left + (width - width * 0.95));
                    TBItem = (int)(top + (height - height * 0.245));


                }
                if (iconSize == 30)
                {

                    LRScroll = (int)(left + (width - width * 0.845));
                    TBSCroll = (int)(top + (height - height * 0.23));
                    LRItem = (int)(left + (width - width * 0.95));
                    TBItem = (int)(top + (height - height * 0.23));


                }

                foreach (Beam b in beams)
                {
                    CreateSetup(LRScroll, TBSCroll, LRItem, TBItem);
                }

            }

            if (width > 1940)
            {
                //MessageBox.Show("Please move External Beam Planning Window to other monitor, sorry!", "Ahhh I can't code for this");

                Form1 form1 = new Form1();
                form1.StartPosition = FormStartPosition.Manual;
                form1.Location = new Point((left + (width / 2 - form1.Width / 2)), (top + (height / 2 - form1.Height / 2)));
                System.Windows.Forms.Application.Run(form1);
            }



            
            app.ClosePatient();
        }

        private static void CreateSetup(int LRScroll, int TBScroll, int LRItem, int TBItem)
        {
            SetCursorPosition(LRScroll, TBScroll);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.RightDown);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.RightUp);

            Thread.Sleep(500);

            keyPress(KeyEventFlags.VK_DOWN);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_DOWN);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_DOWN);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_RETURN);
            Thread.Sleep(20);

            SetCursorPosition(LRItem, TBItem);
            Thread.Sleep(20);

            MouseEvent(MouseEventFlags.LeftDown);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.LeftUp);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.RightDown);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.RightUp);

            Thread.Sleep(500);

            keyPress(KeyEventFlags.VK_UP); //1
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//2
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//3
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//4
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//5
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//6
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//7
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//8
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//9
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//10
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//11
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//12
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_RETURN);
            Thread.Sleep(20);

            SetCursorPosition(LRScroll, TBScroll);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.RightDown);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.RightUp);

            Thread.Sleep(500);

            keyPress(KeyEventFlags.VK_DOWN);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_DOWN);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_DOWN);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_RETURN);
            Thread.Sleep(20);


            SetCursorPosition(LRItem, TBItem);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.LeftDown);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.LeftUp);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_DELETE);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_LEFT);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_RETURN); // one done

            Thread.Sleep(1000);
        }

        private static void CreateSetupLargeResolution( int LRItem, int TBItem)
        {

            SetCursorPosition(LRItem, TBItem);
            Thread.Sleep(20);

            MouseEvent(MouseEventFlags.LeftDown);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.LeftUp);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.RightDown);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.RightUp);

            Thread.Sleep(500);

            keyPress(KeyEventFlags.VK_UP); //1
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//2
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//3
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//4
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//5
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//6
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//7
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//8
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//9
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//10
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//11
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_UP);//12
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_RETURN);
            Thread.Sleep(20);


            SetCursorPosition(LRItem, TBItem);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.LeftDown);
            Thread.Sleep(20);
            MouseEvent(MouseEventFlags.LeftUp);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_DELETE);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_LEFT);
            Thread.Sleep(20);
            keyPress(KeyEventFlags.VK_RETURN); // one done

            Thread.Sleep(1000);
        }

        [Flags]
        public enum MouseEventFlags
        {
            LeftDown = 0x00000002,
            LeftUp = 0x00000004,
            MiddleDown = 0x00000020,
            MiddleUp = 0x00000040,
            Move = 0x00000001,
            Absolute = 0x00008000,
            RightDown = 0x00000008,
            RightUp = 0x00000010
        }

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out MousePoint lpMousePoint);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        public static void SetCursorPosition(int x, int y)
        {
            SetCursorPos(x, y);
        }

        public static void SetCursorPosition(MousePoint point)
        {
            SetCursorPos(point.X, point.Y);
        }

        public static MousePoint GetCursorPosition()
        {
            MousePoint currentMousePoint;
            var gotPoint = GetCursorPos(out currentMousePoint);
            if (!gotPoint) { currentMousePoint = new MousePoint(0, 0); }
            return currentMousePoint;
        }

        public static void MouseEvent(MouseEventFlags value)
        {
            MousePoint position = GetCursorPosition();

            mouse_event
                ((int)value,
                 position.X,
                 position.Y,
                 0,
                 0)
                ;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MousePoint
        {
            public int X;
            public int Y;

            public MousePoint(int x, int y)
            {
                X = x;
                Y = y;
            }
        }

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);




        [Flags]
        public enum KeyEventFlags
        {
            VK_UP = 0x26, //up key
            VK_DOWN = 0x28,  //down key
            VK_LEFT = 0x25,
            VK_RIGHT = 0x27,
            VK_RETURN = 0x0D,
            VK_DELETE = 0x2E,


        }
        public static int keyPress(KeyEventFlags KeyEventFlags)
        {
            const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
            const uint KEYEVENTF_KEYUP = 0x0002;


            //Press the key
            keybd_event((byte)KeyEventFlags, 0, KEYEVENTF_EXTENDEDKEY | 0, 0);
            return 0;
        }

        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(IntPtr classname, string title); // extern method: FindWindow

        [DllImport("user32.dll")]
        static extern void MoveWindow(IntPtr hwnd, int X, int Y,
            int nWidth, int nHeight, bool rePaint);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        static List<IntPtr> GetAllWindows()
        {
            List<IntPtr> windows = new List<IntPtr>();
            EnumWindows(delegate (IntPtr hWnd, IntPtr lParam)
            {
                windows.Add(hWnd);
                return true;
            }, IntPtr.Zero);
            return windows;
        }

        static Tuple<int, int, int, int> WindowTuple(string Id)
        {

            int left = 0;
            int width = 0;
            int top = 0;
            int height = 0;

            var windows = GetAllWindows();
            foreach (var window in windows)
            {
                int length = GetWindowTextLength(window);
                StringBuilder sb = new StringBuilder(length + 1);
                GetWindowText(window, sb, sb.Capacity);
                if (sb.ToString().Contains("External") && sb.ToString().Contains(Id))
                {
                    RECT rect;
                    GetWindowRect(window, out rect);
                    left = rect.Left;
                    width = rect.Right - rect.Left;
                    top = rect.Top;
                    height = rect.Bottom - rect.Top;



                }
            }

            Tuple<int, int, int, int> windowTuple = new Tuple<int, int, int, int>(left, width, top, height);
            return windowTuple;
        }

        // Import the SHAppBarMessage function from the shell32.dll library
        [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int SHAppBarMessage(int dwMessage, ref APPBARDATA data);

        // APPBARDATA structure used by the SHAppBarMessage function
        [StructLayout(LayoutKind.Sequential)]
        private struct APPBARDATA
        {
            public int cbSize;
            public IntPtr hWnd;
            public int uCallbackMessage;
            public int uEdge;
            public RECT rc;
            public bool lParam;
        }

        // RECT structure used by the APPBARDATA structure


        // Constants used by the SHAppBarMessage function
        private const int ABM_GETTASKBARPOS = 5;

    }
}
