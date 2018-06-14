using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace check.win
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("check.win 1.0");
            //foreach (String s in args)
            //{
            //    Console.WriteLine(s);
            //}

            if (args.Contains<String>("-p"))
            {
                int pos = Array.IndexOf(args, "-p");
                Console.WriteLine(args[pos + 1]);
                Console.WriteLine("Reading path ");
            }

            //get path
            //
            //return filepath
            //return filesize
            //return checksum
        }
    }
}
