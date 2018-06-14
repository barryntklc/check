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
        static void Main(string[] args)
        {
            String PROGNAME = "check";
            String OSNAME = "win";
            String VERSION = "1.0";
            String FILETAG = PROGNAME + "." + OSNAME;

            String ABOUT = PROGNAME + "." + OSNAME + " v" + VERSION;

            Console.WriteLine(ABOUT);

            //todo handler for /?

            String SRC_PATH;
            if (args.Contains<String>("/s"))
            {
                int pos = Array.IndexOf(args, "/s");

                SRC_PATH = Path.GetFullPath(args[pos + 1]);

            } else
            {
                SRC_PATH = Path.GetFullPath(".");
            }

            String DEFAULT_LOGNAME = FILETAG + "_" + DateTime.Now.ToString("yyyy-MM-dd_hh-mm-ss") + ".xlsx";
            String LOG_PATH = "";
            if (args.Contains<String>("/log"))
            {
                int pos = Array.IndexOf(args, "/log");

                LOG_PATH = Path.GetFullPath(args[pos + 1]);

                try
                {
                    //https://stackoverflow.com/questions/1395205/better-way-to-check-if-a-path-is-a-file-or-a-directory
                    FileAttributes i = File.GetAttributes(args[pos + 1]);
                    if (i.HasFlag(FileAttributes.Directory))
                    {
                        LOG_PATH = Path.GetFullPath(args[pos + 1]) + "\"" + DEFAULT_LOGNAME;
                    } else
                    {
                        LOG_PATH = args[pos + 1];
                    }
                } catch (Exception e)
                {
                    LOG_PATH = DEFAULT_LOGNAME;
                }
            }
            else
            {
                LOG_PATH = DEFAULT_LOGNAME;
            }

            Console.WriteLine("src_path=\"" + SRC_PATH + "\"");
            Console.WriteLine("dest_path=\"" + LOG_PATH + "\"");

            Console.WriteLine("Loading file listing...");
            FileItem[] FileTree = GetFileTree(SRC_PATH);
            Console.WriteLine("File listing loaded.");

            //Console.WriteLine(LOG_PATH.Equals(""));

            //source:
            //https://tedgustaf.com/blog/2012/create-excel-20072010-spreadsheets-with-c-and-epplus/

            FileInfo EXCEL_PATH = new FileInfo(LOG_PATH);

            Console.WriteLine("Building excel file at \"" + LOG_PATH + "\"");
            BuildExcel(FileTree, EXCEL_PATH);
            Console.WriteLine("Excel file saved.");

            //get path
            //
            //return filepath
            //return filesize
            //return checksum
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
            FileInfo[] fileinfo = new DirectoryInfo(path).GetFiles();
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

                //Console.WriteLine(i.SIZE + ", " + i.CHECKSUM + ", " + i.DEBUG);

                files.Add(i);
            }

            String[] directories = Directory.GetDirectories(path);

            foreach (String d in directories)
            {
                files.AddRange(GetFileTree(d));
            }

            return files.ToArray();
        }

        //source
        //https://stackoverflow.com/questions/1993903/how-do-i-do-a-sha1-file-checksum-in-c/1993910
        static String ComputeSHA1(String filepath)
        {
            using (FileStream fs = new FileStream(filepath, FileMode.Open))
            using (BufferedStream bs = new BufferedStream(fs))
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
