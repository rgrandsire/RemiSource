/*
---------------------------------------------------------------------------------------------------------------
| Date          | Version    | Author             | Comments                                                  |
---------------------------------------------------------------------------------------------------------------
| 08/16/2016    | 1.0.0.1    | Remi G Grandsire   | Original development                                      |
---------------------------------------------------------------------------------------------------------------
| 08/17/2016    | 1.0.0.2   |  "     "   "        | Removed unused stuff                              
| 08/18/2016    | 1.0.0.3   |  "     "   "        | Import complete                                           |
| 09/05/2016    | 1.0.0.4   |  "     "   "        | Fixed issue with unknown records                          |
| 09/06/2016    | 1.0.0.5   |  "     "   "        | Found issue with POPK VS POID                             |
| 09/07/2016    | 1.0.0.6   |  "     "   "        | Issue with multiple popid in the file                     |
| 09/14/2016    | 1.0.0.7   |                     | Get MC DB from Rgistration database                       |
|               |           |                     | Fixed issue with fileds swapped                           |
|               |           |                     | Cleaned up the log file                                   |
| 09/15/2016    | 1.0.0.8   |                     | Use the same receipt number if the POPK does not change   |
|               |           |                     | Added Auto to Receipt number to show Imported VS Manual   |
| 09/16/2016    | 1.0.1.1   |                     | Re-did the whole data import                              |
| 09/19/2016    | 1.0.1.2   |                     | Get the DB prefix from the database (container table)     |
|               |           |                     | Removed login by adding check the debug flag              |
---------------------------------------------------------------------------------------------------------------

ToDo:
- Move the log file path and stuff to the logfile function
- Add the date and time to the login function
- Check the debug flag in the log function 

*/
using System;
using System.IO;
using System.Data; 
using System.Data.SqlClient;
using System.Configuration;

namespace SandvikImport
{
    class Program
    {
        public static string prodDB;
        public static string entusername;
        public static string entpassword;
        public static string rootfilepath;
        public static string importfilepath;
        public static string importarchivefilepath;
        public static string archivesandlogs;
        public static string Logfilepath;
        public static string logfilename;
        public static string LogFilePathAndName;
        public static string DaysToKeepImportFileCopies;
        public static string DaysToKeepLogFiles;
        public static string debugflag;
        public static DateTime iday;
        public static Guid importid = Guid.NewGuid();

        //Start Error Log
        public static errorlogging errorLog = new errorlogging();

        static void Main(string[] args)
        {

            try
            {
                getConfigConstants();
            }
            catch (Exception m)
            {
                Console.WriteLine("There was an error with getting stuff: " + m.Message);
            }
            //Set Constant
            iday = DateTime.Now;


            //SET Logfile Location
            archivesandlogs = (rootfilepath + "\\ArchivesAndLogs\\");
            if (!Directory.Exists(archivesandlogs))
            {
                Directory.CreateDirectory(archivesandlogs);
            }
            Logfilepath = (archivesandlogs + "\\logfiles\\");
            if (!Directory.Exists(Logfilepath))
            {
                Directory.CreateDirectory(Logfilepath);
            }
            LogFilePathAndName = (Logfilepath + logfilename + iday.ToString("MM-dd-yyyy.HHmm") + ".log");
            if (Program.debugflag == "Y")
            {
                errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " ArchivesAndLogs\\logfiles path Set.");
            }


            //Set Paths
            if (Program.debugflag == "Y")
            {
                errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " Setting PartDataImports path.");
            }
            importfilepath = (rootfilepath + "\\PartDataImports\\");
            if (!Directory.Exists(importfilepath))
            {
                Directory.CreateDirectory(importfilepath);
            }
            if (Program.debugflag == "Y")
            {
                errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " Setting ArchivesAndLogs\\ImportedFiles path.");
            }
            importarchivefilepath = (archivesandlogs + "\\ImportedFiles\\");
            if (!Directory.Exists(importarchivefilepath))
            {
                Directory.CreateDirectory(importarchivefilepath);
            }


            //Import Part Data
            errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " Importing Part Information.");
            Console.WriteLine("Importing PO Receipt");


            // Import the file data
            import i = new import();
            i.importPartInfo();

            if (Program.debugflag == "Y")
            {
                errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " Part Import Completed.");
            }

            //Cleanup Old Files
            if (Program.debugflag == "Y")
            {
                errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " Cleaning up old files.");
            }
            cleanup();

            // Export Started
            errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " ********************************** Import Complete **********************************");

            if (Program.debugflag == "Y")
            {
                Console.WriteLine("Press Enter to finish");
                Console.ReadLine();
            }
            try
            {
                Console.Clear();
            }
            catch { }
        }

        public static void getConfigConstants()
        {
            var appSettings = System.Configuration.ConfigurationManager.AppSettings;
            if (appSettings.Count > 0)
            {
                foreach (var key in appSettings.AllKeys)
                {
                    Console.WriteLine("Key: {0} Value: {1}", key, appSettings[key]);
                }
                
                rootfilepath = appSettings["rootfilepath"].ToString() ?? "c:\\temp\\";
                entusername = appSettings["entusername"].ToString() ?? "mczar";
                entpassword = appSettings["entpassword"].ToString() ?? "mczar";
                logfilename = appSettings["logfilename"].ToString() ?? "SandvikImport_";
                DaysToKeepImportFileCopies = appSettings["DaysToKeepImportFileCopies"].ToString() ?? "-10";
                DaysToKeepLogFiles = appSettings["DaysToKeepLogFiles"].ToString() ?? "-10";
                debugflag = appSettings["debug"].ToString() ?? "N";
            }
            
        }
        public static void cleanup()
        {
            //Start Error Log
            errorlogging errorLog = new errorlogging();

            errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " - Cleanup of outdated Log files.");
            string[] xfiles = Directory.GetFiles(Program.Logfilepath);

            int days = Convert.ToInt32(Program.DaysToKeepLogFiles);

            foreach (string file in xfiles)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.LastAccessTime < DateTime.Now.AddDays(days))
                    fi.Delete();
            }
        }
    }
}

