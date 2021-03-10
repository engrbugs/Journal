using System;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Journal
{
    class Program
    {
        const string VERSION = "v0.98";
        const string JOURNAL_PATH = @"C:\Users\engrb\OneDrive\bugs\Journal";

        [DllImport("user32.dll")]
        public static extern int SetForegroundWindow(int hwnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool ShowWindow(IntPtr hWnd, ShowWindowEnum flags);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int SetForegroundWindow(IntPtr hwnd);

        private enum ShowWindowEnum
        {
            Hide = 0,
            ShowNormal = 1, ShowMinimized = 2, ShowMaximized = 3,
            Maximize = 3, ShowNormalNoActivate = 4, Show = 5,
            Minimize = 6, ShowMinNoActivate = 7, ShowNoActivate = 8,
            Restore = 9, ShowDefault = 10, ForceMinimized = 11
        };

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        //This is a replacement for Cursor.Position in WinForms
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        //This simulates a left mouse click
        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            System.Threading.Thread.Sleep(500);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);
        }

        static void BringMainWindowToFront(string processName)
        {
            // get the process
            Process bProcess = Process.GetProcessesByName(processName).FirstOrDefault();

            // check if the process is running
            if (bProcess != null)
            {
                // check if the window is hidden / minimized
                if (bProcess.MainWindowHandle == IntPtr.Zero)
                {
                    // the window is hidden so try to restore it before setting focus.
                    ShowWindow(bProcess.Handle, ShowWindowEnum.Restore);
                }

                // set user the focus to the window
                SetForegroundWindow(bProcess.MainWindowHandle);
            }
            else
            {
                // the process is not running, so start it
                // Process.Start(processName);
            }
        }
        static void Main(string[] args)
        {
            Console.WriteLine($"Creating Journal {VERSION}");
            Console.WriteLine($"Loading WinWord App");
            Word.Application objWord = new Word.Application();
            Console.WriteLine($"Making it visible and maximized");
            objWord.Visible = true;
            objWord.WindowState = Word.WdWindowState.wdWindowStateNormal;

            Console.WriteLine($"Creating a document");
            Word.Document objDoc = objWord.Documents.Add();

            Console.WriteLine($"What is the date for today... hmmm...");
            objWord.Selection.TypeText(strHeader1Text());

            Console.WriteLine($"Typing the header");
            objDoc.Paragraphs[1].set_Style(Word.WdBuiltinStyle.wdStyleHeading1);
            objDoc.Paragraphs[1].Range.Underline = Word.WdUnderline.wdUnderlineSingle;

            Console.WriteLine($"Clear formatting for you to type");
            objWord.Selection.TypeParagraph();
            objWord.Selection.TypeText(Environment.NewLine);
            objDoc.Paragraphs[2].set_Style(Word.WdBuiltinStyle.wdStyleNormal);

            try
            {
                Console.WriteLine($"Making a folder");
                Directory.CreateDirectory($"{JOURNAL_PATH}\\" +
                    $"{strFolderYear()}\\{strFolderMonth()}");
                Console.WriteLine($"I'm close... Saving the file for you");
                objDoc.SaveAs2($"{JOURNAL_PATH}\\{strFolderYear()}\\" +
                    $"{strFolderMonth()}\\{strFilename()}.docx");
            }
            catch (Exception e)
            {
                Console.WriteLine("The process failed: {0}", e.ToString());
            }
            
            Console.WriteLine($"Focus on WinWord app");
            objWord.Activate();
            BringMainWindowToFront("WINWORD.EXE");
            objWord.Visible = true;
            objWord.WindowState = Word.WdWindowState.wdWindowStateMaximize;
            Console.WriteLine($"Bye.");
            Process[] processes = Process.GetProcessesByName("WinWord");
            RECT rct = new RECT();
            foreach (Process p in processes)
            {
                IntPtr windowHandle = p.MainWindowHandle;

                // do something with windowHandle

                
                GetWindowRect(windowHandle, ref rct);

                
            }
            SetCursorPos(972, 1028);
            SetCursorPos(-1950, 800);
            SetCursorPos(-1950, 1240);
            Console.WriteLine($"{rct.Left}, {rct.Top}, {rct.Right}, {rct.Bottom}");
            LeftMouseClick(-1950, 1240);
            Console.WriteLine((int)((rct.Left + rct.Right) / 2));
            Console.WriteLine((int)((rct.Bottom + rct.Top) / 2));
        }

        static string strHeader1Text()
        {
            return $"{DateTime.Now.DayOfWeek}, " +
                $"{DateTime.Now.ToLongDateString()}" +
                $"—{DateTime.Now.ToShortTimeString()}";
        }

        static string strFolderYear()
        {
            return DateTime.Now.Year.ToString();
        }

        static string strFolderMonth()
        {
            return $"{DateTime.Now.Month.ToString("00")}-" +
                $"{DateTime.Now.ToString("MMMM")}";
        }

        static string strFilename()
        {
            return $"{DateTime.Now.ToString("d")}-" +
                $"{DateTime.Now.ToString("HH")}" +
                $"{DateTime.Now.ToString("mm")}";
        }
    }
}
