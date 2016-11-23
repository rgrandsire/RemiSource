/*
---------------------------------------------------------------------------------------------------------------
| Date          | Version    | Author             | Comments                                                  |
---------------------------------------------------------------------------------------------------------------
| 10/11/2016    | 1.0.0.1    | Remi G Grandsire   | Original development                                      |
| 10/12/2016    | 1.0.0.2    |                    | Added configurable offset for file parsing                |
| 10/25/2016    | 1.0.0.3    |                    | Added filename to MC_Interfacelog table                   |
| 10/28/2016    | 1.0.0.4    |                    | Replaced deprecated ConfigurationSettings with            |
|               |            |                    | ConfigurationManager                                      |
| 11/02/2016    | 1.0.0.5    |                    | Remove 3 custom fields from table to use RecordData       |
| 11/18/2016    | 1.0.0.6    |                    | Change to pick any files and delete after import complete |
| 11/21/2016    | 1.0.0.7    |                    | Put log and archive in the Import folder (Archive and Log)|
| 11/24/2016    | 1.0.0.8    |                    | Retrieve the hours from A- in a different column          |
---------------------------------------------------------------------------------------------------------------

*/

using System;
using System.IO;
using System.Data.SqlClient;

namespace SammamishMeterImport
{
    class Program
    {
        // Public stuff used throughout the program
        public static int recUpdated = 0;
        public static int recNotUdated = 0;
        public static string myConStr = "";
        public static string myLogFile = "";
        public static string zDebug = "N";
        public static errorlogging errorLog = new errorlogging();
        public static Guid importid = System.Guid.NewGuid();
        public static DateTime iday;

        static void DBImport(string zSQL)       //This function updates the database with mileage and hours from file
        {
            //Putting data in the database
            SqlConnection newCon = new SqlConnection(myConStr);
            SqlCommand newCom = new SqlCommand(zSQL, newCon);
            int rowsAffected = 0;
            try
            {
                newCon.Open();
                try
                {
                    rowsAffected = newCom.ExecuteNonQuery();
                    if (rowsAffected > 0)
                    {
                        recUpdated++;
                        errorLog.logMessage(myLogFile, rowsAffected.ToString() + " record updated");
                    }
                    else
                    {
                        errorLog.logMessage(myLogFile, "Error updating DB (See statement above)");
                        recNotUdated++;
                    }
                }
                catch (SqlException z)
                {
                    errorLog.logMessage(myLogFile, "SQL Update error: " + z.Message);
                }
                newCon.Close();
            }
            catch (SqlException z)
            {
                errorLog.logMessage(myLogFile, "There was an error adding data to MC DB: " + z.Message);
            }

        }
        static void cleanup(int daysToKeep, string daPath)        // This function is used to clean up log files if needed
        {
            string[] xfiles = Directory.GetFiles(daPath);
            int days = Convert.ToInt32(daysToKeep);
            if (days > 0)
            {
                foreach (string file in xfiles)
                {
                    FileInfo fi = new FileInfo(file);
                    if (fi.LastAccessTime < DateTime.Now.AddDays(-days))
                    {
                        fi.Delete();
                        errorLog.logMessage(myLogFile, "Deleting old file: " + file);
                    }
                }
            }
        }

        static string getMC_DB()
        {
            string connSuccess = "";
            string connKey = System.Configuration.ConfigurationManager.AppSettings["connectionkey"];
            string econnStr = System.Configuration.ConfigurationManager.AppSettings["regdb"];
            string entcontainercode = "";
            string zServer = "";
            string zSql = "SELECT [cr].[dbserver_name], rtrim([c].[container_type_code]) + [c].[Container_Code] FROM [dbo].[Container] [c] INNER JOIN [dbo].[container_resource] [cr] WITH (NOLOCK) ON [c].[container_guid]= [cr].[container_guid] WHERE [cr].[connection_key]=@conKey";
            SqlConnection zCon = new SqlConnection(econnStr);
            SqlCommand newCom = new SqlCommand(zSql, zCon);
            newCom.CommandTimeout = 480;
            newCom.Parameters.Clear();
            newCom.CommandText = zSql;
            newCom.Parameters.AddWithValue("@conKey", connKey);
            if (zDebug == "Y")
            {
                Console.WriteLine(zSql);
            }
            try
            {
                zCon.Open();
                try
                {
                    SqlDataReader myReader = newCom.ExecuteReader();
                    if (myReader.Read())
                    {
                        entcontainercode = myReader[1].ToString();
                        zServer = myReader[0].ToString();
                        connSuccess = entcontainercode + "|" + zServer;
                        if (zDebug == "Y")
                        {
                            errorLog.logMessage(myLogFile, "MC DB: " + entcontainercode);
                            errorLog.logMessage(myLogFile, "MC Server: " + zServer);
                        }
                    }
                    else
                    {
                        connSuccess = "000|000";
                    }
                }
                catch (SqlException v)
                {
                    errorLog.logMessage(myLogFile, "Error getting the entity DB info: " + v.Message);
                    connSuccess = "000|000";
                }
            zCon.Close();
            }
            catch (SqlException z)
            {
                errorLog.logMessage(myLogFile, "There was an error getting the MC DB: " + z.Message);
                Console.WriteLine("Error with sql: " + z.Message);
                connSuccess = "111|111";
                return connSuccess;
            }
            return connSuccess;
        }

        static void FileToMC()
        {
            //Run Stored Procedure to import odometer readings
            try
            {
                SqlConnection newCon = new SqlConnection(myConStr);
                SqlCommand cmd = new SqlCommand();
                Int32 rowsAffected;
                cmd.CommandText = "CSTM_SammamishMetersToMC";
                cmd.CommandType = System.Data.CommandType.StoredProcedure;
                cmd.CommandTimeout = 480;
                cmd.Connection = newCon;
                newCon.Open();
                rowsAffected = cmd.ExecuteNonQuery();
                newCon.Close();
                errorLog.logMessage(myLogFile, "Import Stored Procedure completed.");
            }
            catch (SqlException e)
            {
                errorLog.logMessage(myLogFile, "Import Stored Procedure error: [" + e.ErrorCode + "] " + e.Message);
            }
        }

        static void moveProcessedFile(string fname, string zPath)
        {
            // Move the file to the processed folder
            string filenameDate = iday.ToString("MM-dd-yyyy.HHmm");
            ////////////////// something is wrong
            string sourceFile = (zPath +"\\"+ fname);
            string destinationFile = (zPath+ "\\Archive\\" + fname);
            System.IO.File.Move(sourceFile, destinationFile + "." + filenameDate);
            errorLog.logMessage(myLogFile, fname + " has been archived and renamed to: " + destinationFile + "." + filenameDate);
        }

        static void Main(string[] args)
        {
            Console.Title = "MC meter import utility";
            iday = DateTime.Now;
            //Console.
            //Let's check if path existss
            string zFile = System.Configuration.ConfigurationManager.AppSettings["ImportFilePath"];
            if (!Directory.Exists(zFile+"/Archive"))
            {
                Directory.CreateDirectory(zFile+"/Archive");
            }
            if (!Directory.Exists(zFile + "/Log"))
            {
                Directory.CreateDirectory(zFile + "/LOG");
            }
            myLogFile = zFile+"/Log/SammamishMeterImport_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
            errorLog.logMessage(myLogFile, "Data extraction tool started");
            int days4Files = Convert.ToInt16(System.Configuration.ConfigurationManager.AppSettings["DaysToKeep"]);
            // Get the connection stuff
            string MCUserName = System.Configuration.ConfigurationManager.AppSettings["entusername"];
            string MCPassword = System.Configuration.ConfigurationManager.AppSettings["entpassword"];
            
            string DBSql = "";
            string odometerReading = "";
            string hourReading = "";
            string zVehicle = "";
            int i = 0;
            int zGood = 0;
            int zBad = 0;
            int VehicleOffset = Convert.ToInt16(System.Configuration.ConfigurationManager.AppSettings["VehicleOffset"]);
            int MilesOffset = Convert.ToInt16(System.Configuration.ConfigurationManager.AppSettings["MilesOffset"]);
            int HoursOffset = Convert.ToInt16(System.Configuration.ConfigurationManager.AppSettings["HoursOffset"]);
            zDebug = System.Configuration.ConfigurationManager.AppSettings["Debug"];
            var utcNow = DateTime.UtcNow;
            string[] arr1 = new string[] { "", "" };
            if (zDebug == "Y")
            {
                errorLog.logMessage(myLogFile, "MC User name: " + MCUserName);
                errorLog.logMessage(myLogFile, "MC User P/W: " + MCPassword);
            }
            DBSql = getMC_DB();
            arr1 = DBSql.Split('|');
            myConStr = "Server=" + arr1[1] + ";Database=" + arr1[0] + ";Integrated Security=False; User ID=" + MCUserName + "; Password=" + MCPassword;
            if (zDebug == "Y")
            {
                errorLog.logMessage(myLogFile, "MC connection string: " + myConStr);
            }
            Console.WriteLine("");
            errorLog.logMessage(myLogFile, "Serial #                    | ID                 | Hours         | Mileage");
            errorLog.logMessage(myLogFile, "###########################################################################");
            Console.WriteLine("");
            Console.WriteLine("Retrieving data one vehicle at the time");
            string[] files = Directory.GetFiles(zFile);
            //get file count - if "0" then email sig
            int filecount = files.Length;
            if (filecount != 0)
            {
                // Loop through files and import
                foreach (string file in files)
                {
                    FileInfo fi = new FileInfo(file);
                    string fname = fi.Name;

                    if (zDebug == "Y")
                    {
                        errorLog.logMessage(myLogFile, " Processing file: " + fname);
                    }
                    if (File.Exists(file))
                        foreach (string line in File.ReadLines(file))
                            try
                            {
                                i++;
                                string[] words = line.Split(',');
                                Console.Write(".");
                                odometerReading = words[MilesOffset].Trim();
                                //hourReading = words[HoursOffset].Trim();
                                hourReading = words[HoursOffset].Trim(new Char[] { ' ', '*', '"' });
                                //Console.WriteLine("Hours: " + hourReading);
                                if (hourReading.Length > 2)
                                {
                                    string[] arr2 = hourReading.Split('-');
                                    if (arr2[0] == "A")
                                    {
                                        hourReading = arr2[1];
                                    }
                                    else hourReading = "0";
                                }
                                else hourReading = "0";
                                zVehicle = words[VehicleOffset].Trim(new Char[] { ' ', '*', '"' });
                                errorLog.logMessage(myLogFile, zVehicle + "\t\t|" + odometerReading + "\t\t|" + hourReading);
                                DBSql = "insert into MC_InterfaceLog (ImportID, FileName, RecordData, RecordNumber) Values ('" +
                                        importid + "','" + fname + "', '" + hourReading + "," + odometerReading + "," +
                                        zVehicle + "','" + (i).ToString() + "');";
                                if (zDebug == "Y")
                                {
                                    errorLog.logMessage(myLogFile, "Query: " + DBSql);
                                }
                                if (zVehicle.Length > 0)
                                {
                                    zGood++;
                                    DBImport(DBSql);
                                }
                                else zBad++;
                                // Execute the stored procedure to load the data to the asset table
                                FileToMC();

                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("Error with import: " + e.Message);
                            }
                    Console.WriteLine("");
                    Console.WriteLine(zGood.ToString() + " records imported in the database");
                    Console.WriteLine(zBad.ToString() + " records could not be imported");
                    Console.WriteLine("Retrieved all records and inserted them to the database");
                    errorLog.logMessage(myLogFile, "###########################################################################");
                    errorLog.logMessage(myLogFile, "");
                    errorLog.logMessage(myLogFile, recUpdated.ToString() + " successfully updated");
                    errorLog.logMessage(myLogFile, recNotUdated.ToString() + " could not be updated");
                    errorLog.logMessage(myLogFile, "Import complete, check above for success or failure");
                    if (zDebug == "Y")
                    {
                        Console.WriteLine("Press any key to end this program");
                        Console.ReadKey();
                    }

                    errorLog.logMessage(myLogFile, " Move processed file.");
                    moveProcessedFile(fname, zFile);
                cleanup(days4Files, zFile+"\\Archive\\");
                cleanup(days4Files, zFile+"\\Log\\");
                }
            }
        }
    }
}

