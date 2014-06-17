using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;

namespace MonitorFolder
{
    class MonitorFolder
    {
        //make a choice if the program continue to monitor the folder
        public static bool flag = true;
        //Show nofify message
        string message = "";
        public static DateTime currentTime;
        FileSystemWatcher watcher;

        //Start to monitor the specificed folder
        public void Run(ArrayList monitorPath) 
        {
            foreach (string path in monitorPath)
            {
                 watcher = new FileSystemWatcher();
                 watcher.Path = path;
                 watcher.IncludeSubdirectories = false;

                 //Watch for changes in LastAccess and LastWrite times, and the renaming of files or directories. 
                 //watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;

                 //Add event handlers
                 // watcher.Changed += new FileSystemEventHandler(OnChanged);
                 watcher.Created += new FileSystemEventHandler(OnCreated);
                 watcher.Deleted += new FileSystemEventHandler(OnDeleted);
                 watcher.Renamed += new RenamedEventHandler(OnRenamed);        

                 //Begain watching.
                 watcher.EnableRaisingEvents = true;
            }
            Thread.Sleep(120000);
            while(Notify.flag);
        }

        //Show notify dialog
        private void ShowNotify(string message,string path,DateTime time) 
        {
            Notify notify = new Notify(message + "  Are you sure to exit this program and not monitor the folder again?",path,time);
            notify.ShowDialog();
        }

        #region Event Handle
        private void OnDeleted(object sender, FileSystemEventArgs e)
        {
            currentTime = DateTime.Now;
            message = "Folder  "+e.Name+"  Deleted!";
            ShowNotify(message, e.FullPath, currentTime);
        }

        private void OnRenamed(object sender, RenamedEventArgs e)
        {
            currentTime = DateTime.Now;
            message = "Folder  " + e.Name + "  Renamed!";
            ShowNotify(message, e.FullPath, currentTime);
        }

        private void OnCreated(object sender, FileSystemEventArgs e)
        {
            currentTime = DateTime.Now;
            message = "Folder  " + e.Name + "  Created!";
            ShowNotify(message, e.FullPath, currentTime);
        }

        private void OnChanged(object sender, FileSystemEventArgs e)
        {
            currentTime = DateTime.Now;
            message = "Folder  " + e.Name + "  Created!";
            ShowNotify(message, e.FullPath, currentTime);
        }     
        #endregion
    }
}
