using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace MonitorFolder
{
    class Program
    {
        static bool flag = true;
        static string input = "";
        static MonitorFolder monitorFolder;

        static void Main(string[] args)
        {
            ArrayList arrayList = new ArrayList();
            //arrayList.Add(@"C:\Users\v-sihe\Desktop\Test\Test1");
            //arrayList.Add(@"C:\Users\v-sihe\Desktop\Test\Test2");
            //arrayList.Add(@"C:\Users\v-sihe\Desktop\Test\Test3");
            //arrayList.Add(@"\\iisdist\release\oob\Antares\hosting");
            //arrayList.Add(@"\\iisdist\release\oob\Antares\hosting_release");
            //arrayList.Add(@"\\iisdist\release\oob\Antares\rd_websites_n");
            arrayList.Add(@"\\reddog\builds\branches\rd_websites_n_release");

            monitorFolder = new MonitorFolder();

            while (flag)
            {
                Console.WriteLine("Please input a command like:  monitor/gettime/exit");

                input = Console.ReadLine();

                if (input.Equals("monitor"))
                {
                    monitorFolder.Run(arrayList);
                }
                else if (input.Equals("gettime"))
                {
                    while (input != "exit")
                    {
                        Console.WriteLine("Please input the time when this bug opened(just like 2012/1/1):");
                        input = Console.ReadLine();
                        new GetTime(Convert.ToDateTime(input)).GetSpanDays();

                        Console.WriteLine("Back to Parent directory please input \"exit\":");
                        input = Console.ReadLine();
                        if (input.Equals("exit"))
                        {
                            continue;
                        }
                    }
                }
                else if (input.Equals("exit"))
                {
                    flag = false;
                }
                else
                {
                    Console.WriteLine("Unknow input, please try again!");
                }
                Console.WriteLine("Please press any key to continue:");
                Console.ReadKey();
            }
        }
    }
}
