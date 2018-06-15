using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace check.win
{
    class Program
    {
        static Boolean VERBOSE = false;

        static String PROGNAME = "check";
        static String OSNAME = "win";
        static String VERSION = "1.0";
        static String FILETAG = PROGNAME + "." + OSNAME;

        static void Main(string[] args)
        {
            if (args.Contains<String>("/?"))
            {
                String HELP_MSG = @"Dumps the path, size, and checksum of files into an Excel spreadsheet (.xlsx).

By default, scans the current folder and dumps information to a spreadsheet in the same folder.

USAGE:
    check.win.exe /s [source directory path] | [ /? | 
                                                /v | 
                                                /verbose | 
                                                /log [excel file path] ]

Options:
    /?                          Display this help message
    /v  /verbose                Show extra information, for diagnostic purposes
    /s [source directory path]  Check all files from the [source directory path]
    /log [excel file path]      Save the file to the specified [excel file path]
";

                Console.WriteLine(HELP_MSG);
            }
            else
            {
                String ABOUT = PROGNAME + "." + OSNAME + " v" + VERSION;
                Console.WriteLine(ABOUT);

                String SRC_PATH;
                if (args.Contains<String>("/s"))
                {
                    int arg_position = Array.IndexOf(args, "/s");

                    SRC_PATH = Path.GetFullPath(args[arg_position + 1]);

                }
                else
                {
                    SRC_PATH = Path.GetFullPath(".");
                }

                String DEFAULT_LOGNAME = FILETAG + "_" + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss") + ".xlsx";
                String LOG_PATH = "";
                if (args.Contains<String>("/log"))
                {
                    int arg_position = Array.IndexOf(args, "/log");

                    LOG_PATH = Path.GetFullPath(args[arg_position + 1]);

                    try
                    {
                        //https://stackoverflow.com/questions/1395205/better-way-to-check-if-a-path-is-a-file-or-a-directory
                        FileAttributes i = File.GetAttributes(args[arg_position + 1]);
                        if (i.HasFlag(FileAttributes.Directory))
                        {
                            LOG_PATH = Path.GetFullPath(args[arg_position + 1]) + "\\" + DEFAULT_LOGNAME;
                        }
                        else
                        {
                            LOG_PATH = args[arg_position + 1];
                        }
                    }
                    catch (Exception e)
                    {
                        LOG_PATH = args[arg_position + 1];
                    }
                }
                else
                {
                    LOG_PATH = DEFAULT_LOGNAME;
                }

                if (args.Contains<String>("/v") || args.Contains<String>("/verbose"))
                {
                    VERBOSE = true;
                }

                if (VERBOSE)
                {
                    Console.WriteLine("src_path=\"" + SRC_PATH + "\"");
                    Console.WriteLine("log_path=\"" + LOG_PATH + "\"");
                    Console.WriteLine("verbose=" + VERBOSE);
                }

                Console.WriteLine("Loading file listing...");
                FileItem[] FileTree = GetFileTree(SRC_PATH);
                Console.WriteLine("File listing loaded.");

                //source:
                //https://tedgustaf.com/blog/2012/create-excel-20072010-spreadsheets-with-c-and-epplus/

                FileInfo EXCEL_PATH = new FileInfo(LOG_PATH);

                Console.WriteLine("Building excel file.");
                BuildExcel(FileTree, EXCEL_PATH);
                Console.WriteLine("Excel file saved to \"" + LOG_PATH + "\".");
            }
        }

        //https://tedgustaf.com/blog/2012/create-excel-20072010-spreadsheets-with-c-and-epplus/
        static void BuildExcel(FileItem[] filelisting, FileInfo filename)
        {
            ExcelPackage package = new ExcelPackage(filename);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Output");

            int row = 1;
            foreach (FileItem s in filelisting)
            {
                worksheet.Cells[row, 1].Value = s.PATH;
                worksheet.Cells[row, 2].Value = s.SIZE;
                worksheet.Cells[row, 3].Value = s.CHECKSUM;
                worksheet.Cells[row, 4].Value = s.DEBUG;

                row++;
            }

            //worksheet.Column(1).Width = 40;
            worksheet.Column(1).AutoFit();
            worksheet.Column(2).AutoFit();
            worksheet.Column(3).AutoFit();
            worksheet.Column(4).AutoFit();

            package.Save();
        }

        static FileItem[] GetFileTree(String path)
        {
            List<FileItem> files = new List<FileItem>();
            FileInfo[] fileinfo = new DirectoryInfo(path).GetFiles(); //throws exceptions if path too long
            foreach (FileInfo f in fileinfo)
            {
                FileItem i = new FileItem();
                i.PATH = f.FullName;
                i.SIZE = f.Length.ToString();

                try
                {
                    i.CHECKSUM = ComputeSHA1(f.FullName).Replace("-", "");
                }
                catch (Exception e)
                {
                    i.DEBUG = e.Message;
                }

                if (VERBOSE)
                {
                    Console.WriteLine(i.PATH);
                    Console.WriteLine("\t" + i.SIZE + ", " + i.CHECKSUM + ", " + i.DEBUG);
                }

                files.Add(i);
            }

            String[] directories = Directory.GetDirectories(path);

            foreach (String d in directories)
            {
                try
                {
                    files.AddRange(GetFileTree(d)); //too long
                }
                catch (Exception e) //catch folder exceptions
                {
                    Console.WriteLine(d);
                    FileItem i = new FileItem();
                    i.PATH = d;
                    i.DEBUG = e.Message;

                    files.Add(i);
                }
            }

            return files.ToArray();
        }

        //source
        //https://stackoverflow.com/questions/1993903/how-do-i-do-a-sha1-file-checksum-in-c/1993910
        static String ComputeSHA1(String filepath)
        {
            using (FileStream fs = new FileStream(filepath, FileMode.Open)) //check for unnecessary write requests
            using (BufferedStream bs = new BufferedStream(fs)) //check for
            {
                using (SHA1Managed sha1 = new SHA1Managed())
                {
                    String hash = BitConverter.ToString(sha1.ComputeHash(bs));
                    return hash;
                }
            }
        }

        class FileItem
        {
            public String PATH { get; set; }
            public String SIZE { get; set; }
            public String CHECKSUM { get; set; }
            public String DEBUG { get; set; }
        }
    }
}
