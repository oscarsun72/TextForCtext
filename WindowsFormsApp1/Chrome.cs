using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace TextForCtext
{
    public class Chrome
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, string lpString, int nMaxCount);

        //static void Main(string[] args)
        //{
            static IntPtr handle = GetForegroundWindow();
            static string title = GetActiveTabTitle(handle);

            //Console.WriteLine("Active tab title: " + title);
            //Console.ReadLine();
        //}
        internal static string ActiveTabTitle { get { return title; } }
        internal static string GetActiveTabTitle(IntPtr handle)
        {
            const int nChars = 256;
            string buffer = new string(' ', nChars);
            GetWindowText(handle, buffer, nChars);

            string title = buffer.Trim();
            if (title.EndsWith("Google Chrome"))
            {
                int tabIndex = title.LastIndexOf(" - Google Chrome");
                title = title.Substring(0, tabIndex);
            }
            return title;
        }


    }
}
