using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace CheckPathAfterCtrlC
{
    public partial class MonitorCtrlC : Form
    {
        int count = 0;
        public MonitorCtrlC()
        {
            InitializeComponent();

            Clipboard.Clear();

            KeyboardHook kh = new KeyboardHook();
            kh.SetHook();
            kh.OnKeyDownEvent += kh_OnKeyDownEvent;
        }
     
        private void kh_OnKeyDownEvent(object sender, KeyEventArgs e)
        {
            //if (e.KeyData == (Keys.S | Keys.Control)) { MessageBox.Show("Ctrl + S"); }
            //if (e.KeyData == (Keys.H | Keys.Control)) { MessageBox.Show("Ctrl + H"); }
            //if (e.KeyData == (Keys.A | Keys.Control | Keys.Alt)) { MessageBox.Show("Ctrl + Alt + A"); }
            //if (e.KeyData == (Keys.C | Keys.Control)) { MessageBox.Show("Ctrl + C"); }

            //Note: The hook will exec before the system response, that means you need press Ctrl + C two times
            if (e.KeyData == (Keys.C | Keys.Control))
            {
                getTextDataFromClipboard();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            (new KeyboardHook()).UnHook();
        }

        private void getTextDataFromClipboard() 
        {
            //When the count is 0, the clipboard have nothing this time due to clear() method in init step
            if (count > 0)
            {
                IDataObject iData = Clipboard.GetDataObject();
                if (iData.GetDataPresent(DataFormats.Text))
                {
                    TextInClipborard.Text = (String)iData.GetData(DataFormats.Text);
                    isPathExist();
                }
                else
                {
                    TextInClipborard.Text = "Could not retrieve text data from the clipboard!";
                    TextInClipborard.BackColor = Color.Yellow;
                    Message.Text = "The data in clipboard is not Text format, please try a valid path!";
                }
            }
            
            count++;
        }

        private void isPathExist() 
        {
            string path = TextInClipborard.Text.Trim();

            if (File.Exists(path) || Directory.Exists(path))
            {
                TextInClipborard.BackColor = Color.Green;
                Message.Text = "The path is valid!";
            }
            else 
            {
                TextInClipborard.BackColor = Color.Red;
                Message.Text = "The path is invalid or you have no permission to access it!";
            }
        }
    }   
}
