using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;

namespace SandvikImport
{
    class import
    {
        errorlogging errorLog = new errorlogging();   // Next time put the date time stuff in the object rather than the call
        string ImportFilePath = Program.importfilepath;
        string ImportFileArchivePath = Program.importarchivefilepath;
        string LogFilePathAndName = Program.LogFilePathAndName;
        string zServer = ""; 
        string zUser = Program.entusername;
        string zPassword = Program.entpassword;
        string prodDB = "";
        
        static string getMC_DB()
        { /////////Getting the ent database and sql server information from the Registration database
            string connSuccess = "111|000";
            string connKey = System.Configuration.ConfigurationSettings.AppSettings["connectionkey"] ?? "Sandvik";
            string econnStr = System.Configuration.ConfigurationSettings.AppSettings["regdb"] ?? "mcRegistrationSA";
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
            catch(SqlException z)
            {
                Console.WriteLine("Error with sql: " + z.Message);
                connSuccess = "111|111";
                Console.WriteLine("3: " + connSuccess);
                if (zCon.State == ConnectionState.Open)
                    zCon.Close();
                return connSuccess;
            }
           
                
        }
        public void importPartInfo()
        {
            string[] files = Directory.GetFiles(ImportFilePath);
            //get file count - if "0" then email sig
            int filecount = files.Length;
            if (filecount != 0)
            {
                // Loop through files and import
                foreach (string file in files)
                {
                    FileInfo fi = new FileInfo(file);
                    string fname = fi.Name;

                    if (Program.debugflag == "Y")
                    {
                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Processing file: " + fname + ".");
                    }
                    ProcessPartInformation(fname);
                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Move processed file.");
                    moveProcessedFile(fname);

                }
            }
            errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " - Cleanup old files.");
            cleanup();
        }
        public void ProcessPartInformation(string importFileName)
        {
            string[] arr1 = new string[] { "", "" };
            int counter = 0;
            string zPOPK = "";
            int zErrNum = 0;
            string sql = "";
            prodDB = getMC_DB();
            arr1 = prodDB.Split('|');
            prodDB = arr1[0];
            zServer = arr1[1];
            bool zFlag = false;
           int jj = 1;
            bool newRecord = true;
            bool firstRun = true;
            bool noPOPK = false;
            if (Program.debugflag == "Y")
            {
                Console.WriteLine("entServer Server: " + zServer);
                Console.WriteLine("ent DB: " + prodDB);
            }
                SqlConnection conn = new SqlConnection("Data Source=" + zServer + ";Initial Catalog=" + prodDB + ";Integrated Security=False; User ID=" + zUser + "; Password=" + zPassword);
            
            SqlCommand cmd = new SqlCommand(sql, conn);
            try
            {
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
            }
            catch (Exception er)
            {
                Console.WriteLine("Error with DB: " + er);
                errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + "Error with DB: " + er);
            }
            string zPKInvoice = "";
            string zPK = "";
            foreach (string line in File.ReadLines(ImportFilePath + importFileName))
            {
                counter++;
                if ((counter != 1) && (line.Length > 1))// skip first line
                {
                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " =============================================================================================================================================================================================================================");
                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Record #: " + (counter - 1).ToString());
                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Data: " + line);
                    string[] words = line.Split('|');
                    zPOPK = (words[0]).Trim();///////////////// That's where I need to check if the POPK changes...
                   // if (oldPOID == zPOPK) zFlag = false; else zFlag = true;
                    string zLineItemNum = words[8].Trim();
                    string zUnitPrice = words[6].Trim();
                    int zRows = 0;
                    //Getting the POPK from the POID
                    if (Program.debugflag == "Y")
                    {
                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Getting the POPK from POID: " + zPOPK);
                    }
                    sql = "select POPK from PurchaseOrder where POID=" + zPOPK;
                    cmd.CommandTimeout = 480;
                    cmd.CommandType = CommandType.Text;
                    cmd.Parameters.Clear();
                    cmd.CommandText = sql;
                    if (Program.debugflag == "Y")
                    {
                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Query= " + sql);
                    }
                    try
                    {
                        SqlDataReader reader = cmd.ExecuteReader();
                        if (reader.Read())
                        {
                            zPOPK = reader.GetValue(0).ToString();   // There is an existing PO in the DB  --> I need to check if it's a new record or update the receipt name
                            reader.Close();
                            noPOPK = false;
                            ///////////// Now checking if it's a new record
                            cmd.CommandText = "Select count(*) from PurchaseOrderInvoice  WITH (NOLOCK) Where popk = " + zPOPK + " and Type='R';";
                            if (Program.debugflag == "Y")
                            {
                                errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Checking for new record, SQL= " + cmd.CommandText);
                            }
                            try
                            {
                                int count = (int)cmd.ExecuteScalar();
                                jj = count + 1;
                                if (Program.debugflag == "Y")
                                {
                                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Number of receipt lines= " + count.ToString());
                                }
                                if (count == 0)
                                    newRecord = true;
                                else
                                    newRecord = false;
                            }
                            catch (SqlException q)
                            {
                                errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " There was an error querying for new record: " + q.Message);
                                errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Record : " + line + " could not be processed");
                            }
                            if (((newRecord) && (zFlag)) || (firstRun))
                            {
                                // I need to insert the receipt and receipt number into PurchaseOrderInvoice
                                firstRun = false;
                                sql = "insert into PurchaseOrderInvoice (POPK, Subtotal, FreightCharge, TaxAmount, Total, ReceiptNo, ReceiptDate, Type, TypeDesc, ReceiptNoInternal) values " +     //  sql = "insert into PurchaseOrderInvoice (POPK, Subtotal, FreightCharge, TaxAmount, Total, ReceiptNo, ReceiptDate, Type) values " +
                                    "(@POPK, 0, 0, 0, @Total, @Receipt,Convert(varchar,Getdate(), 101), 'R', 'Receipt',@ReceiptNoInternal)";
                                cmd.Parameters.Clear();
                                cmd.CommandText = sql;
                                cmd.Parameters.Clear();
                                cmd.Parameters.Add("@Total", SqlDbType.Float);
                                cmd.Parameters["@Total"].Value = Convert.ToDouble("44.55");
                                cmd.Parameters.Add("@POPK", SqlDbType.Int);
                                cmd.Parameters["@POPK"].Value = zPOPK;
                                cmd.Parameters.Add("@ReceiptNoInternal", SqlDbType.Int);
                                cmd.Parameters["@ReceiptNoInternal"].Value = jj.ToString();
                                cmd.Parameters.Add("@Receipt", SqlDbType.VarChar);
                                cmd.Parameters["@Receipt"].Value = zPOPK + "-" + jj.ToString(); // ---> replace zLine num with Recordnointernal
                                if (Program.debugflag == "Y")
                                {
                                    Console.WriteLine("SQL1= " + cmd.CommandText.ToString());
                                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Insert record in PurchaseOrderInvoice");
                                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL: " + sql);
                                }
                                ///////////// Let's execute this SQL thing and see...
                                try
                                {
                                    zRows = cmd.ExecuteNonQuery();
                                    if (Program.debugflag == "Y")
                                    {
                                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Row(s) affected= " + zRows.ToString());
                                    }
                                    //////////// Let's get the PK and the PKInvoice
                                    reader.Close();
                                }
                                catch (SqlException s)
                                {
                                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " There was an error inserting the record in the POI table: " + s.Message + "-" + s.ErrorCode);
                                    errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Switching to next record");
                                }
                            }
                            else jj--;
                            sql = "select zPKInvoice= Convert(Varchar,PurchaseOrderInvoice.InvoicePK) from PurchaseOrderInvoice where POPK=" + zPOPK + " and ReceiptNoInternal= " + jj.ToString() + ";";
                            cmd.Parameters.Clear();
                            cmd.CommandType = CommandType.Text;
                            Console.WriteLine("PKInvoice query: " + sql);
                            cmd.CommandText = sql;
                            if (Program.debugflag == "Y")
                            {
                                errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Getting the PKInvoice for the POPK (" + zPOPK + ")");
                            }
                            try {
                                reader = cmd.ExecuteReader();
                                if (reader.Read())
                                {
                                    zPKInvoice = reader.GetString(0);
                                    if (Program.debugflag == "Y")
                                    {
                                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " PKInvoice= " + zPKInvoice);
                                    }
                                    /// From there I need to insert the record into the POIdetail table
                                    cmd.Parameters.Clear();
                                    sql = "select zPK= Convert(Varchar,PurchaseOrderDetail.PK) from PurchaseOrderDetail where POPK=" + zPOPK + " and LineItemNo=" + zLineItemNum + "; ";
                                    cmd.CommandText = sql;
                                    if (Program.debugflag == "Y")
                                    {
                                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Retrieving the PK: " + sql);
                                    }
                                    reader.Close();
                                    try
                                    {
                                        reader = cmd.ExecuteReader();
                                        if (reader.Read())
                                        {
                                            zPK = reader.GetString(0);
                                            if (Program.debugflag == "Y")
                                            {
                                                errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " PK= " + zPK);
                                            }
                                        }
                                    } 
                                    catch (SqlException r)
                                    {
                                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SqlError: " + r.Message);
                                        zErrNum++;
                                    }
                                    ////////////////////////////
                                    sql = "insert into PurchaseOrderInvoiceDetail(InvoicePK, PurchaseOrderDetailPK, OrderUnitQtyReceived, OrderUnitQtyBackOrdered, OrderUnitQtyCanceled, Bin) values " +
                                         "(" + zPKInvoice + ", " + zPK + ", " + words[10] + ", 0, 0, (select Bin from PurchaseOrderDetail where POPK=" + zPOPK + " and LineItemNo=" + zLineItemNum + "));";
                                    cmd.Parameters.Clear();
                                    cmd.CommandText = sql;
                                    reader.Close();
                                    if (Program.debugflag == "Y")
                                    {
                                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL4: " + sql);
                                    }
                                    try
                                    {
                                        cmd.ExecuteNonQuery();
                                    }
                                    catch (SqlException a)
                                    {
                                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL Error: " + a.Message);
                                        zErrNum++;
                                    }
                                    //////////////////// I need to update the price in 2 tables
                                    sql = "update Part set LastOrderUnitPrice=" + zUnitPrice + " where PartID= '" + words[4].Trim() + "';";
                                    cmd.Parameters.Clear();
                                    cmd.CommandText = sql;
                                    if (Program.debugflag == "Y")
                                    {
                                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL5: " + sql);
                                    }
                                    try
                                    {
                                        cmd.ExecuteNonQuery();
                                    }
                                    catch (SqlException x)
                                    {
                                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Error updating Part table: " + x.Message + "//" + x.ErrorCode);
                                    }
                                    sql = "update PurchaseOrderDetail set OrderUnitPrice=" + zUnitPrice + " where popk= " + zPOPK + " and LineItemNo=" + zLineItemNum + ";";
                                    cmd.Parameters.Clear();
                                    cmd.CommandText = sql;
                                    if (Program.debugflag == "Y")
                                    {
                                        errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL6: " + sql);
                                    }
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            catch (SqlException e)
                            {
                                errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL Error: " + e.Message);
                                errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL: " + sql);
                                zErrNum++;
                            }
                        }
                        else
                        {
                            errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Could not find the PO: " + zPOPK + " in the database, skip to next record");
                            noPOPK = true;
                            reader.Close();
                        }

                    }
                    catch (SqlException v)
                    {
                        if (Program.debugflag == "Y")
                        {
                            errorLog.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " There was error getting the PK Invoice: " + v.Message);
                        }
                        zErrNum++;
                    }
                }

                if ((zErrNum > 0) && (noPOPK== false))
                {
                    errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + zErrNum.ToString() + " errors were reported");
                    errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " ");
                    errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " - " + importFileName + " Error while processing - Rolling back changes");
                    cmd.CommandText = "delete from PurchaseOrderInvoice where POPK=" + zPOPK + ";";
                    
                    cmd.ExecuteNonQuery();
                }
            }
            if (conn.State == ConnectionState.Open)
            { 
                conn.Close();
            }
            errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + "---------------------------------------------------------------------------------------------------------------------------------------------------");
            errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " " + importFileName + " imported, " + (counter-1).ToString() + " rows in file.");
            counter++;
       }


        public void moveProcessedFile(string fname)
        {
            // Move the file to the processed folder
            string filenameDate = Program.iday.ToString("MM-dd-yyyy.HHmm");

            string sourceFile = (ImportFilePath + fname);
            string destinationFile = (ImportFileArchivePath + fname);
            System.IO.File.Move(sourceFile, destinationFile + "." + filenameDate);
            errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " " + fname + " has been archived and renamed to: " + destinationFile + "." + filenameDate);
        }

        public void cleanup()
        {
            errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " Cleaning up outdated Import files.");
            string[] xfiles = Directory.GetFiles(ImportFileArchivePath);

            int days = Convert.ToInt32(Program.DaysToKeepImportFileCopies);

            foreach (string file in xfiles)
            {
                FileInfo fi = new FileInfo(file);
                if (fi.LastAccessTime < DateTime.Now.AddDays(days))
                {
                    fi.Delete();
                    errorLog.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " " + fi.Name + " has been purged");
                }
            }
        }

    }
}

