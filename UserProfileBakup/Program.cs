using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using uparchiver;

namespace uparchiver
{
    class Program
    {
        static void Main(string[] args)
        {
            string source = "ngd005-caz";
            string dest = @"\\mgmt01-caz\user_profiles$";

            machinearchive ma = new machinearchive();

            ma.archivelocation = dest;
            ma.machineName = source;
            ma.userFolderList = new string[] { "robert.hogarth"};

            Console.WriteLine("Starting Archive");
            long byteCount = ma.Archive();

            Console.WriteLine("Completed Archive - " + byteCount.ToString() + " Bytes written");
        }
    }
}
