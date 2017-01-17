/*
---------------------------------------------------------------------------------------------------------------
| Date          | Version    | Author             | Comments                                                  |
---------------------------------------------------------------------------------------------------------------
| 09/06/2016    | 1.0.0.1    | Remi G Grandsire   | Original development                                      |
---------------------------------------------------------------------------------------------------------------
| 09/07/2016    | 1.0.0.2   |                     | Added the database connection and log maintenance         |
| 09/14/2016    | 1.0.0.3   |                     | Get the database name from Reg DB                         |
|               |           |                     | Get the settings from appconfig instead of ini            |
| 09/19/2016    | 1.0.0.4   |                     | Change the query to get the Entity name                   |
|               |           |                     | Use the debug flag to write to log                        |
| 09/21/2016    | 1.0.1.1   |                     | Use InterfaceLog table and stored procedure to load data  |
| 11/14/2016    | 1.0.1.2   |                     | Calculate Meter to Miles and seconds to hours             |
---------------------------------------------------------------------------------------------------------------
*/


using System;
using System.IO;
using System.Collections.Generic;
using System.Data.SqlClient;
using Geotab.Checkmate.ObjectModel;
using Geotab.Checkmate.ObjectModel.Engine;
using Geotab.Checkmate;

namespace NavistarImport
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


        static void cleanup(int daysToKeep)        // This function is used to clean up log files if needed
        {
            string zFolder = Path.GetDirectoryName(myLogFile);
            string[] xfiles = Directory.GetFiles(zFolder);
            int days = Convert.ToInt32(daysToKeep);
             if (days >0) {
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
                ////// Get a reader object to get the data........
                
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
        static void DBImport(string zSQL)       //This function updates the database with mileage and hours from Geotab
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
        static void GeoTabToMC()
        {
            //Run Stored Procedure to import odometer readings
            try
            {
                SqlConnection newCon = new SqlConnection(myConStr);
                SqlCommand cmd = new SqlCommand();
                Int32 rowsAffected;
                cmd.CommandText = "CSTM_Navistar_GeoTabToMC";
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
        static void Main(string[] args)
        {
            Console.Title = "MC meter and hour import utility";
            //Console.
            //Let's check if path exists
            if (!Directory.Exists("C:\\temp\\"))
            {
                Directory.CreateDirectory("C:\\temp\\");
            }
            myLogFile = "C:\\temp\\GeoToNavistar_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
            errorLog.logMessage(myLogFile, "Data extraction tool started");
            int days4Files =  Convert.ToInt16(System.Configuration.ConfigurationManager.AppSettings["DaysToKeep"]);
                        // Get the connection stuff
            string geoUserName = System.Configuration.ConfigurationManager.AppSettings["GeotabUser"];
            string geoPassWord = System.Configuration.ConfigurationManager.AppSettings["GeotabPassWord"]; 
            string geoServer = System.Configuration.ConfigurationManager.AppSettings["GeotabServer"];
            string geoDB = System.Configuration.ConfigurationManager.AppSettings["GeotabDB"];
            string MCUserName = System.Configuration.ConfigurationManager.AppSettings["entusername"];
            //string MCDB = "";  //  --> get from registration database
            string MCPassword = System.Configuration.ConfigurationManager.AppSettings["entpassword"];
            string DBSql = "";
            string odometerReading = "";
            string hourReading = "";
            string zVehicle = "";
            zDebug = System.Configuration.ConfigurationManager.AppSettings["Debug"];
            var utcNow = DateTime.UtcNow;
            string[] arr1 = new string[] { "", "" };
            // Let's cleanup the log folder before doing anything
            cleanup(days4Files);
            if (zDebug == "Y")
            {
                errorLog.logMessage(myLogFile, "MC User name: " + MCUserName);
                errorLog.logMessage(myLogFile, "MC User P/W: " + MCPassword);
            }
            DBSql = getMC_DB();
            arr1 = DBSql.Split('|');
            ///////////////"Server=localhost; Integrated Security=False; Database=mcregistrationSA; User Id=mczar; Password=mczar" />
            myConStr =  "Server="+ arr1[1]+";Database=" + arr1[0] + ";Integrated Security=False; User ID=" + MCUserName + "; Password=" + MCPassword;
            if (zDebug == "Y")
            {
                errorLog.logMessage(myLogFile, "MC connection string: " + myConStr);
            }
            //Connecting to server and starting the API
            try
            {
                var api = new API(geoUserName, geoPassWord, null, geoDB, geoServer);
                Console.WriteLine("Connected succesfully to GeoTab");
                Console.WriteLine("");
                Console.WriteLine("Getting all devices and count");
                var devices = api.Call<IList<Device>>("Get", typeof(Device));
                errorLog.logMessage(myLogFile, "Found " + devices.Count + " vehicles to import data for");
                Console.WriteLine("");
                errorLog.logMessage(myLogFile, "Serial #                    | ID                 | Hours         | Mileage");
                errorLog.logMessage(myLogFile, "###########################################################################");
                Console.WriteLine("Found " + devices.Count + " vehicles to import");
                Console.WriteLine("");
                Console.WriteLine("Retrieving data one vehicle at the time");
                for (int i = 0; i < devices.Count; i++)
                {                    
                    Device device = devices[i];
                    Console.Write(".");
                    // Search for status data based on the current device and the odometer reading
                    var statusMilesSearch = new StatusDataSearch();
                    var statusHoursSearch = new StatusDataSearch();
                    statusMilesSearch.DeviceSearch = new DeviceSearch(device.Id);
                    statusHoursSearch.DeviceSearch = new DeviceSearch(device.Id);
                    statusMilesSearch.DiagnosticSearch = new DiagnosticSearch(KnownId.DiagnosticOdometerAdjustmentId);
                    statusHoursSearch.DiagnosticSearch = new DiagnosticSearch(KnownId.DiagnosticEngineHoursAdjustmentId);
                    statusMilesSearch.FromDate = DateTime.MaxValue;
                    statusHoursSearch.FromDate = DateTime.MaxValue;
                    // Retrieve the odometer status data
                    IList<StatusData> statusMiles = api.Call<IList<StatusData>>("Get", typeof(StatusData), new { search = statusMilesSearch });
                    IList<StatusData> statusHours = api.Call<IList<StatusData>>("Get", typeof(StatusData), new { search = statusHoursSearch });
                    double zMilesToday = Convert.ToDouble(statusMiles[0].Data * 0.000621371);
                    double zTimeToday = Convert.ToDouble(statusHours[0].Data / 3600);
                    odometerReading = (Math.Round(zMilesToday)).ToString();
                    hourReading = (Math.Round(zTimeToday,2)).ToString();
                    zVehicle =  device.ToString();
                    string[] words = zVehicle.Split(':');
                    zVehicle = words[0].Substring(0, words[0].Length - 2);
                    errorLog.logMessage(myLogFile, zVehicle + "|" + odometerReading.ToString() + "|" + hourReading.ToString());
                    //Load the stuff in the database (MC_InterfaceLog table now)
                    //DBSql = "update Asset set Meter1Reading=" + odometerReading.ToString() + ", Meter2Reading=" + hourReading.ToString() + " where AssetID='" + zVehicle.Trim() + "';";
                    string zData = hourReading.ToString()+"|" + odometerReading.ToString()+ "|" + zVehicle.Trim();
                    DBSql = "insert into MC_InterfaceLog with (Rowlock) (Hours, Miles,  VehicleID, ImportID, RecordData, RecordNumber) Values ('" + 
                            hourReading.ToString()+"','"+odometerReading.ToString() +"','" + zVehicle.Trim() + "','"+ importid+"', '"+ zData + "','"+ (i+1).ToString()+"');";
                    if (zDebug == "Y")
                    {
                        errorLog.logMessage(myLogFile, "Query: " + DBSql);
                    }
                    DBImport(DBSql);
                }
                // Execute the stored procedure to load the data to the asset table
                //GeoTabToMC();
                Console.WriteLine();
                Console.WriteLine("Retrieved all records and inserted them to the database");
            }
            catch (Exception e)
            {
                Console.WriteLine("Error connecting to GeoTab: " + e.Message);
            }
            errorLog.logMessage(myLogFile, "###########################################################################");
            errorLog.logMessage(myLogFile, "");
            cleanup(days4Files);
            errorLog.logMessage(myLogFile, recUpdated.ToString() + " successfully updated");
            errorLog.logMessage(myLogFile, recNotUdated.ToString() + " could not be updated");
            errorLog.logMessage(myLogFile, "Import complete, check above for success or failure");
            if (zDebug == "Y")
            {
                Console.WriteLine("Press any key to end this program");
                Console.ReadKey();
            }
         }
     }
}

