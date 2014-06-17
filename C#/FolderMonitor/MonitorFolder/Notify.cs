using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MonitorFolder
{
    public partial class Notify : Form
    {
        public static bool flag = true;
        DateTime time;

        public Notify(string message,string path,DateTime time)
        {
            InitializeComponent();
            timer1.Interval = 1000;
            timer1.Start();

            showMessage.Text = message;
            LinkToPath.Text = path;
            UpdatedTime.Text = "Updated Time:  "+time.ToString();

            this.TopMost = true;
            this.OK.Focus();
        }

        private void Cancle_Click(object sender, EventArgs e)
        {
            flag = false;
            this.Close();
        }

        private void OK_Click(object sender, EventArgs e)
        {
            flag = true;
            this.Close();
        }

        private void OK_GotFocus(object sender, System.EventArgs e)
        {
            FlashIt();
        }

         #region FlashIcon
         private void FlashIt() 
        { 
            FLASHWINFO fi = new FLASHWINFO(); 
            fi.cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf(fi); 
            fi.hwnd = Handle; 
            fi.dwFlags = FLASHW_TRAY; 
            fi.uCount = 3; 
            fi.dwTimeout = 0; 
            FlashWindowEx(ref fi);
        } 

        [DllImport("user32.dll")] 
        [return: MarshalAs(UnmanagedType.Bool)] 
        static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

        [StructLayout(LayoutKind.Sequential)] 
        public struct FLASHWINFO 
        { 
            public UInt32 cbSize; 
            public IntPtr hwnd; 
            public UInt32 dwFlags; 
            public UInt32 uCount; 
            public UInt32 dwTimeout;
        } 

        //Stop flashing. The system restores the window to its original state. 
        //public const UInt32 FLASHW_STOP = 0; 
        //Flash the window caption. 
        public const UInt32 FLASHW_CAPTION = 1; 
        //Flash the taskbar button. 
        public const UInt32 FLASHW_TRAY = 2; 
        //Flash both the window caption and taskbar button. 
        //This is equivalent to setting the FLASHW_CAPTION | FLASHW_TRAY flags. 
        public const UInt32 FLASHW_ALL = 3; 
        //Flash continuously, until the FLASHW_STOP flag is set. 
        public const UInt32 FLASHW_TIMER = 4; 
        //Flash continuously until the window comes to the foreground. 
        public const UInt32 FLASHW_TIMERNOFG = 12;

        #endregion

        private void LinkToPath_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkToPath.LinkVisited = true;
            System.Diagnostics.Process.Start(LinkToPath.Text);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            time = DateTime.Now;
            TimeSpan timeSpan = time - MonitorFolder.currentTime;
            currentTime.Text = "Time Span:  "+timeSpan.ToString();
        }
    }
}
