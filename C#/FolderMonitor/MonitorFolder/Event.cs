using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonitorFolder
{
    class Event
    {
        public delegate void FlashIconEventHandler(object sender,EventArgs e);

        public event FlashIconEventHandler FlashIcon;

        public void StartFlashIcon() 
        {
            if (this.FlashIcon != null) 
            {
                this.FlashIcon(this,new EventArgs());
            }
        }
    }
}
