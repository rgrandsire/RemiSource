using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;


namespace SandvikImport
{
    class import
    {
        static errorlogging errorLogNew = new errorlogging();   // Next time put the date time stuff in the object rather than the call
        public static string ImportFilePath = Program.importfilepath;
        public static string ImportFileArchivePath = Program.importarchivefilepath;
        string LogFilePathAndName = Program.LogFilePathAndName;
        string zServer = ""; 
        string zUser = Program.entusername;
        string zPassword = Program.entpassword;
        string prodDB = "";
                

        
        public void importPOInfo()
        {
            extraStuff.getTheFiles();   // sFtp transfer 
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
                        errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Processing file: " + fname + ".");
                    }
                    ProcessPOInformation(fname);
                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Move processed file.");
                    extraStuff.moveProcessedFile(fname);

                }
            }
            errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " - Cleanup old files.");
            extraStuff.cleanup();
        }
        public void ProcessPOInformation(string importFileName)
        {

            string[] arr1 = new string[] { "", "" };        // Stores the server name and the ent database
            int counter = 0;                                // Use to display he record number in the log file
            string zPOPK = "";                              // POPK used throughout the program
            string zPOID = "";                              // POID from the file used to get the POPK
            int zErrNum = 0;                                // Error flag: if not equal to 0 then it's a rool back
            string sql = "";
            prodDB = extraStuff.getMC_DB();                 // String used to get the server name and the ent db into the array
            arr1 = prodDB.Split('|');
            prodDB = arr1[0];
            zServer = arr1[1];
            bool zFlag = false;                             // Not used anymore
           int jj = 1;                                      // Receipt number appended to PO number for the poinvoicedetails lines
            bool newRecord = true;                          // Whether the PO already has receipts
            bool firstRun = true;                           // Whether it is the first record of a file
            bool noPOPK = false;                            // POPK validation
            if (Program.debugflag == "Y")
            {
                Console.WriteLine("entServer Server: " + zServer);
                Console.WriteLine("ent DB: " + prodDB);
            }
            SqlConnection conn = new SqlConnection("Data Source=" + zServer + ";Initial Catalog=" + prodDB + "; Integrated Security=False; User ID=" + zUser + "; Password=" + zPassword);
            SqlCommand cmd = new SqlCommand(sql, conn);
            SqlCommand cmd2 = new SqlCommand(sql, conn);
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
                errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + "Error with DB: " + er);
            }

            ////////////////////////////////////////////////////////////////// The program should not execute if we cannot connect to the database

            string zPKInvoice = "";
            string zPK = "";
            foreach (string line in File.ReadLines(ImportFilePath + importFileName))
            {
                noPOPK = true;
                counter++;
                zPOPK = "";
                zPOID = "";
                zPK = "";
                zPKInvoice = "";
                if ((counter != 1) && (line.Length > 1))// skip first line
                {
                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " =============================================================================================================================================================================================================================");
                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Record #: " + (counter - 1).ToString());
                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Data: " + line);
                    string[] words = line.Split('|');
                    zPOID = (words[0]).Trim();///////////////// That's where I need to check if the POPK changes...
                    string zLineItemNum = "";
                    string zPartID = words[5].Trim();
                    string zPartName = words[4].Trim();
                    string zUnitPrice = words[6].Trim();
                    string zTotalPrice = words[1].Trim();
                    string zVendorName = words[3].Trim();
                    int zRows = 0;
                    //Getting the POPK from the POID
                    if (Program.debugflag == "Y")
                    {
                        errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Getting the POPK from POID: " + zPOPK);
                    }
                    try
                    {
                        cmd2.Parameters.Clear();
                        cmd2.CommandTimeout = 480;
                        sql = "select isnull((select POPK from PurchaseOrder WITH (NOLOCK) where POID=" + zPOID + " and VendorName= '" + zVendorName + "'),'0');";    // Adding the vendor name for 3 way matching
                        cmd2.CommandType = CommandType.Text;
                        cmd2.CommandText = sql;
                        if (Program.debugflag == "Y")
                        {
                            errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Query= " + sql);
                        }
                        using (SqlDataReader POreader1 = cmd2.ExecuteReader())
                        {
                            if (POreader1.Read() || POreader1 !=null)
                            {
                                zPOPK = POreader1.GetValue(0).ToString();// There is an existing PO in the DB  --> I need to check if it's a new record or update the receipt name
                                if (zPOPK == "0")
                                {
                                    noPOPK = true;
                                }
                                else
                                {
                                    noPOPK = false;
                                }

                                if (Program.debugflag == "Y")
                                {
                                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " POPK: " + zPOPK);
                                }
                            } 
                            else
                            {
                                zPOPK = "0";
                                // ErrorCode:001 - Could not get the POPK No insert possible
                                errorlogging.errorReport(zPOID, line, 1);
                            }
                            POreader1.Close();
                        }
                        if (zPOPK=="0")
                        {
                            noPOPK = true;
                            errorlogging.errorReport(zPOID, line, 1);
                        }
                    }
                    catch (SqlException v)
                    {
                        errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " There was error getting the POPK : " + v.Message);
                        zErrNum++;
                    }

                    if (!noPOPK)
                    {
                        ///////////// Now checking if it's a new record
                        cmd2.CommandText = "Select count(*) from PurchaseOrderInvoice  WITH (NOLOCK) Where popk = " + zPOPK + " and Type='R';";
                        if (Program.debugflag == "Y")
                        {
                            errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Checking for new record, SQL= " + cmd2.CommandText);
                        }
                        try
                        {
                            int count = (int)cmd2.ExecuteScalar();
                            jj = count + 1;
                            if (Program.debugflag == "Y")
                            {
                                errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Number of receipt lines= " + count.ToString());
                            }
                            if (count == 0)
                                newRecord = true;
                            else
                                newRecord = false;
                        }
                        catch (SqlException q)
                        {
                            errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " There was an error querying for new record: " + q.Message);
                            errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Record : " + line + " could not be processed");
                        }
                        if (((newRecord) && (zFlag)) || (firstRun))
                        {
                            // I need to insert the receipt and receipt number into PurchaseOrderInvoice
                            firstRun = false;
                            sql = "insert into PurchaseOrderInvoice (POPK, Subtotal, FreightCharge, TaxAmount, Total, ReceiptNo, ReceiptDate, Type, TypeDesc, ReceiptNoInternal) values " +
                                "(@POPK, 0, 0, 0, @Total, @Receipt,Convert(varchar,Getdate(), 101), 'R', 'Receipt',@ReceiptNoInternal)";
                            cmd.Parameters.Clear();
                            cmd.CommandText = sql;
                            cmd.Parameters.Clear();
                            cmd.Parameters.Add("@Total", SqlDbType.Float);
                            cmd.Parameters["@Total"].Value = Convert.ToDouble(zTotalPrice);
                            cmd.Parameters.Add("@POPK", SqlDbType.Int);
                            cmd.Parameters["@POPK"].Value = zPOPK;
                            cmd.Parameters.Add("@ReceiptNoInternal", SqlDbType.Int);
                            cmd.Parameters["@ReceiptNoInternal"].Value = jj.ToString();
                            cmd.Parameters.Add("@Receipt", SqlDbType.VarChar);
                            cmd.Parameters["@Receipt"].Value = zPOPK + "-" + jj.ToString();
                            if (Program.debugflag == "Y")
                            {
                                Console.WriteLine("SQL1= " + cmd.CommandText.ToString());
                                errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Insert record in PurchaseOrderInvoice");
                                errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL: " + sql);
                            }
                            ///////////// Let's execute this SQL thing and see...
                            try
                            {
                                zRows = cmd.ExecuteNonQuery();
                                if (Program.debugflag == "Y")
                                {
                                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Row(s) affected= " + zRows.ToString());
                                }

                            }
                            catch (SqlException s)
                            {
                                errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " There was an error inserting the record in the POI table: " + s.Message + "-" + s.ErrorCode);
                                errorlogging.errorReport(zPOID, line, 2);
                                errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Switching to next record");
                            }
                        }
                        else jj--;
                        sql = "select isnull((select zPKInvoice= PurchaseOrderInvoice.InvoicePK from PurchaseOrderInvoice with (nolock) where PurchaseOrderInvoice.POPK=" + zPOPK + " and ReceiptNoInternal='" + jj.ToString() + "'),'999');";
                        cmd2.Parameters.Clear();
                        cmd2.CommandType = CommandType.Text;
                        Console.WriteLine("PKInvoice query: " + sql);
                        cmd2.CommandText = sql;
                        if (Program.debugflag == "Y")
                        {
                            errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Getting the PKInvoice for the POPK (" + zPOPK + ")");
                        }
                        try
                        {
                            using (SqlDataReader reader1 = cmd2.ExecuteReader())
                            {
                                if (reader1.Read() || reader1 != null)
                                {
                                    zPKInvoice = reader1.GetValue(0).ToString();
                                    reader1.Close();
                                    if (Program.debugflag == "Y")
                                    {
                                        errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL: " + sql);
                                        errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " PKInvoice= " + zPKInvoice);
                                    }
                                }
                            }
                        }
                        catch (SqlException e)
                        {
                            errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL Error: " + e.Message);
                            errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL: " + sql);
                            errorlogging.errorReport(zPOID, line, 3);
                            zErrNum++;
                        }
                        /// From there I need to insert the record into the POIdetail table
                        cmd.Parameters.Clear();
                        //Building the SQL
                        sql="if (select count(*)  from PurchaseOrderDetail where PurchaseOrderDetail.POPK = "+zPOPK+" and PurchaseOrderDetail.PartID = '"+zPartID+"') > 0 ";
                        sql += " begin ";
                        sql += " select zPK = Convert(Varchar, PurchaseOrderDetail.PK), zLine = LineItemNo from PurchaseOrderDetail where PurchaseOrderDetail.POPK = " + zPOPK + " and PurchaseOrderDetail.PartID = '" + zPartID + "';";
                        sql += " end else ";
                        sql += " begin ";
                        sql += " select zPK = '999',zLine = 999 ";
                        sql += " end;";
                        cmd.CommandText = sql;
                            if (Program.debugflag == "Y")
                                {
                                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Retrieving the PK: " + sql);
                                }
                                    zLineItemNum = "999";
                            try
                            {
                                using (SqlDataReader reader2 = cmd.ExecuteReader())
                                {
                                    if (reader2.Read() || reader2 != null)
                                    {
                                    zLineItemNum = reader2.GetInt32(1).ToString();
                                    zPK = reader2.GetString(0);
                                    if (Program.debugflag == "Y")
                                    {
                                        errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " LineItemNumber: " + zLineItemNum);
                                        errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " PK= " + zPK);
                                    }
                                    if (zPKInvoice=="999")
                                    {
                                        zErrNum++;
                                        errorlogging.errorReport(zPOID, line, 4);
                                    }
                                    if (zLineItemNum == "999")
                                    {
                                        zErrNum++;
                                        errorlogging.errorReport(zPOID, line, 4);
                                    }
                                }
                                else
                                    {
                                    if (Program.debugflag == "Y")
                                        {
                                            errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Error retrieving the PK and/or line number");
                                        }
                                    zErrNum++;
                                    errorlogging.errorReport(zPOID, line, 4);
                                    }
                                }
                            }
                            catch (SqlException r)
                            {
                                errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SqlError: " + r.Message);
                                zErrNum++;
                            }
                            if (zErrNum == 0)
                            {
                                sql = "insert into PurchaseOrderInvoiceDetail(InvoicePK, PurchaseOrderDetailPK, OrderUnitQtyReceived, OrderUnitQtyBackOrdered, OrderUnitQtyCanceled, Bin) values " +
                                      "(" + zPKInvoice + ", " + zPK + ", " + words[10] + ", 0, 0, (select Bin from PurchaseOrderDetail where POPK=" + zPOPK + " and LineItemNo=" + zLineItemNum + "));";
                                cmd.Parameters.Clear();
                                cmd.CommandText = sql;
                                if (Program.debugflag == "Y")
                                {
                                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL4: " + sql);
                                }
                                try
                                {
                                    cmd.ExecuteNonQuery();
                                }
                                catch (SqlException a)
                                {
                                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL Error: " + a.Message);
                                    zErrNum++;
                                }
                                //////////////////// I need to update the price in 2 tables
                                sql = "update Part set LastOrderUnitPrice=" + zUnitPrice + " where PartID= '" + zPartID + "';";
                                cmd.Parameters.Clear();
                                cmd.CommandText = sql;
                                if (Program.debugflag == "Y")
                                {
                                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL5: " + sql);
                                }
                                try
                                {
                                    cmd.ExecuteNonQuery();
                                }
                                catch (SqlException x)
                                {
                                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " Error updating Part table: " + x.Message + "//" + x.ErrorCode);
                                }
                                sql = "update PurchaseOrderDetail set OrderUnitPrice=" + zUnitPrice + " where popk= " + zPOPK + " and LineItemNo=" + zLineItemNum + ";";
                                cmd.Parameters.Clear();
                                cmd.CommandText = sql;
                                if (Program.debugflag == "Y")
                                {
                                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " SQL6: " + sql);
                                }
                                cmd.ExecuteNonQuery();
                            }
                       
                    }
                }
                else
                {
                    errorLogNew.logMessage(Program.LogFilePathAndName, DateTime.Now.ToString() + " No POPK found for POID:" + zPOID);             
                }

                if ((zErrNum > 0) && (noPOPK== false))
                {
                    errorLogNew.logMessage(LogFilePathAndName, DateTime.Now.ToString() + zErrNum.ToString() + " errors were reported");
                    errorLogNew.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " ");
                    errorLogNew.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " - " + importFileName + " Error while processing - Rolling back changes");
                    zErrNum = 0;
                    cmd.Parameters.Clear();
                    cmd.CommandText = "delete from PurchaseOrderInvoice where POPK=" + zPOPK + ";";
                    cmd.ExecuteNonQuery();
                }
            }
            if (conn.State == ConnectionState.Open)
            { 
                conn.Close();
            }
            errorLogNew.logMessage(LogFilePathAndName, DateTime.Now.ToString() + "---------------------------------------------------------------------------------------------------------------------------------------------------");
            errorLogNew.logMessage(LogFilePathAndName, DateTime.Now.ToString() + " " + importFileName + " imported, " + (counter-1).ToString() + " rows in file.");
            counter++; 
       }
    }
}

