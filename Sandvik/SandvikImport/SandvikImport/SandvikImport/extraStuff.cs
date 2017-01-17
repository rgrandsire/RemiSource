using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using Renci.SshNet;


namespace SandvikImport
{
    public class extraStuff
    {
        static errorlogging errorLog = new errorlogging();


        //MultipleActiveResultSets=True; --> Just in case I need to use it later

        public static string getMC_DB()
        { /////////Getting the ent database and sql server information from the Registration database
            string connSuccess = "111|000";
            string connKey = System.Configuration.ConfigurationManager.AppSettings["connectionkey"] ?? "Sandvik";
            string econnStr = System.Configuration.ConfigurationManager.AppSettings["regdb"] ?? "mcRegistrationSA";
            string entcontainercode = "";
            string zServer = "";
            string zSql = "SELECT [cr].[dbserver_name], rtrim([c].[container_type_code]) + [c].[Container_Code] FROM [dbo].[Container] [c] INNER JOIN [dbo].[container_resource] [cr] WITH (NOLOCK) ON [c].[container_guid]= [cr].[container_guid] WHERE [cr].[connection_key]=@conKey";
            SqlConnection zCon = new SqlConnection(econnStr);
            SqlCommand newCom = new SqlCommand(zSql, zCon);
            newCom.CommandTimeout = 480;
            newCom.Parameters.Clear();
            newCom.CommandText = zSql;
            newCom.Parameters.AddWithValue("@conKey", connKey);

            try
            {
                zCon.Open();
                SqlDataReader myReader = newCom.ExecuteReader();
                ////// Get a reader object to get the data........
                if (myReader.Read())
                {
                    entcontainercode = myReader[1].ToString();
                    zServer = myReader[0].ToString();
                    connSuccess = entcontainercode + "|" + zServer;
                    Console.WriteLine("1: " + connSuccess);
                    if (zCon.State == ConnectionState.Open)
                        zCon.Close();
                    return connSuccess;   //return both the ent DB name and SQL server in the same string to be parsed
                }
                else
                {
                    connSuccess = "000|000";
                    Console.WriteLine("2: " + connSuccess);
                    if (zCon.State == ConnectionState.Open)
                        zCon.Close();
                    return connSuccess;
                }

            }
            catch (SqlException z)
            {
                Console.WriteLine("Error with sql: " + z.Message);
                connSuccess = "111|111";
                Console.WriteLine("3: " + connSuccess);
                if (zCon.State == ConnectionState.Open)
                    zCon.Close();
                return connSuccess;
            }


        }


        public static void getTheFiles()
        {
            string UName = System.Configuration.ConfigurationManager.AppSettings["sFtpUser"] ?? "sandvik";
            string PWord = System.Configuration.ConfigurationManager.AppSettings["sFtpPassword"] ?? "4kK0@o9I";
            string zHost = System.Configuration.ConfigurationManager.AppSettings["sFtpHost"] ?? "ftp2.verian.net";
            string remoteDirectory = System.Configuration.ConfigurationManager.AppSettings["sFtpRemotePath"] ?? "/production/MT_ReceivingPO/"; ////
            string localDirectory = @Program.importfilepath;

            using (var sftp = new SftpClient(zHost, UName, PWord))
            {
                try
                {
                    sftp.Connect();
                }

                catch (Exception aa)
                {
                    Console.WriteLine("Error connecting to sFTP server " + aa.Message);
                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Error with sFtpServer: " + aa.Message);
                }
                var files = sftp.ListDirectory(remoteDirectory);

                foreach (var file in files)
                {
                    string remoteFileName = remoteDirectory + file.Name;
                    string localFileName = localDirectory + file.Name;
                    if ((!file.Name.StartsWith(".")) && !extraStuff.checkForDupe(file.Name))

                        using (Stream file1 = File.OpenWrite(localFileName))
                        {
                            sftp.DownloadFile(remoteFileName, file1);
                        }
                    Console.WriteLine("new File: " + file.Name + " saved in: " + localDirectory);
                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Getting new file: " + file.Name + " saved as: " + localDirectory);
                }
                sftp.Disconnect();
            }

        }
        public static bool checkForDupe(string zName)
        {
            string prodDB = getMC_DB();
            string[] arr1 = prodDB.Split('|');
            string dupe = "";
            string sql = "SELECT '1' FROM MC_InterfaceLog WITH(nolock) WHERE fileName='" + zName + "';";
            string zServer = arr1[1];
            prodDB = arr1[0];
            Console.WriteLine("Server: " + zServer);
            Console.WriteLine("Database: " + prodDB);
            SqlConnection conn1 = new SqlConnection("Data Source=" + zServer + ";Initial Catalog=" + prodDB + ";Integrated Security=False; User ID=mczar; Password=mczar; MultipleActiveResultSets=True;");
            SqlCommand cmd = new SqlCommand(sql, conn1);
            cmd.CommandTimeout = 480;
            cmd.Parameters.Clear();
            cmd.CommandText = sql;
            try
            {
                if (conn1.State == ConnectionState.Closed)
                {
                    conn1.Open();
                    Console.WriteLine("Connecting to MC DB to check for duplicate files");
                }
                ////////// Let's write the code for dupe check
                SqlDataReader aReader = cmd.ExecuteReader();
                ////// Get a reader object to get the data........
                if (aReader.Read())
                {
                    dupe = aReader.GetString(0);
                    Console.WriteLine("Got stuff: " + dupe);
                    aReader.Close();
                    sql = "insert into MC_InterfaceLog (ProcessDate, Filename, ErrorMessage, Processed) values (GetDate(), '" + zName + "', 'File already processed', 'N');";
                    if (Program.debugflag == "Y")
                    {
                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + "SQL: " + sql);
                    }
                    cmd.CommandText = sql;
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (SqlException zEx)
                    {
                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + "Sql Error: " + zEx.Message);
                    }
                    return true;
                }
                else
                {
                    Console.WriteLine("Got nothing no Dupe");
                    // I need to write the file name to the DB
                    sql = "insert into MC_InterfaceLog (ProcessDate, Filename, ErrorMessage, Processed) values (GetDate(), '" + zName + "', 'File ready to be processed', 'Y');";
                    cmd.CommandText = sql;
                    if (Program.debugflag == "Y")
                    {
                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + "SQL: " + sql);
                    }
                    aReader.Close();
                    cmd.ExecuteNonQuery();
                    return false;
                }

            }
            catch (Exception er)
            {
                Console.WriteLine("Error with DB: " + er);
                errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + "Error with DB***: " + er);
            }
            if (conn1.State == ConnectionState.Open)
            {
                conn1.Close();
            }
            return true;
        }

        public static void cleanup()
        {
            //Start Error Log
            errorlogging errorLog = new errorlogging();

            errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " - Cleanup of outdated Log files.");
            string[] xfiles = Directory.GetFiles(Program.Logfilepath);

            int days = Convert.ToInt32(Program.DaysToKeepLogFiles);

            foreach (string file in xfiles)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.LastAccessTime < DateTime.Now.AddDays(days))
                    fi.Delete();
            }
        }

        public static void moveProcessedFile(string fname)
        {
            // Move the file to the processed folder
            string filenameDate = Program.iday.ToString("MM-dd-yyyy.HHmm");

            string sourceFile = (import.ImportFilePath + fname);
            string destinationFile = (import.ImportFileArchivePath + fname);
            System.IO.File.Move(sourceFile, destinationFile + "." + filenameDate);
            errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " " + fname + " has been archived and renamed to: " + destinationFile + "." + filenameDate);
        }
    }
}
