using clibSync.com;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace consoleSync
{
    class Program
    {
        static void Main(string[] args)
        {
            SyncLibCOM syncLib = new SyncLibCOM();
            ConsoleKeyInfo mykey;
            do
            {
                syncLib.FireEvent("Event fired at:" + DateTime.Now.ToString());

                Console.WriteLine("Press q to close application");
                mykey = Console.ReadKey();

            } while (mykey.KeyChar != 'q');
        }
    }
}
