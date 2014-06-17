using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MonitorFolder
{
    class GetTime
    {
        //bool flag = true;

        DateTime today = DateTime.Now;
        DateTime FileBugday;

        public GetTime(DateTime time) 
        {
            FileBugday = time;
        }

        public void GetSpanDays() 
        {
            TimeSpan interval = today - FileBugday;
            Console.WriteLine("Today is:  "+today);
            Console.WriteLine("File bug day is:  "+FileBugday);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("The span time is:  "+interval.Days+"  days");
            Console.ForegroundColor = ConsoleColor.White;
        }
    }
}
