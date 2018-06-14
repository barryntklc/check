using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace check.win
{
    class Program
    {
        static void Main(string[] args)
        {
            String PROGNAME = "check";
            String OSNAME = "win";
            String VERSION = "1.0";
            String FILETAG = PROGNAME + "." + OSNAME;

            String about = PROGNAME + "." + OSNAME + " v" + VERSION;

            Console.WriteLine(about);
            //foreach (String s in args)
            //{
            //    Console.WriteLine(s);
            //}

            String path;
            if (args.Contains<String>("-p"))
            {
                int pos = Array.IndexOf(args, "-p");

                path = Path.GetFullPath(args[pos + 1]);

            } else
            {
                path = Path.GetFullPath(".");
            }
            Console.WriteLine("path=\"" + path + "\"");

            Console.WriteLine("Loading file listing...");
            String[] filelisting = GetFileListing(path);

            Console.WriteLine("You have the following files:");
            foreach (String s in filelisting)
            {
                Console.WriteLine(s);
            }

            //source:
            //https://tedgustaf.com/blog/2012/create-excel-20072010-spreadsheets-with-c-and-epplus/
            String DefaultOutputName = FILETAG + "_" + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss") + ".xlsx";

            Console.WriteLine(DefaultOutputName);

            //get path
            //
            //return filepath
            //return filesize
            //return checksum
        }

        //https://tedgustaf.com/blog/2012/create-excel-20072010-spreadsheets-with-c-and-epplus/
        static void BuildExcel(FileInfo filename)
        {
            ExcelPackage package = new ExcelPackage(filename);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Output");


        }

        static String[] GetFileListing(String path)
        {
            List<String> files = Directory.GetFiles(path).ToList<String>();

            String[] directories = Directory.GetDirectories(path);

            foreach (String d in directories)
            {
                files.AddRange(GetFileListing(d));
            }

            return files.ToArray();
        }
    }
}
