using System;
using System.IO;
using System.Data;
using System.Data.SqlClient;


namespace SandvikImport
{
  class errorlogging
  {

     public void logMessage(string LogFilePathAndName, string message)
    {
      if (!File.Exists(LogFilePathAndName))
      {
        // Create a file to write to.
        StreamWriter swNew = File.CreateText(LogFilePathAndName);
        swNew.WriteLine(message);
        swNew.Close();
      }
      else
      {
        StreamWriter swAppend = File.AppendText(LogFilePathAndName);
        swAppend.WriteLine(message);
        swAppend.Close();
      }
    }
     public static void errorReport(string zPONum, string zData, int ErrCode)
    {
            // Trying to determine what error do we have and what needs to be done in order to fix it
            // The error code is used ti figure out at waht part of the import the error occured
            /*
             * For instance
             * 1: First SQL statement --> PO does not exist or vendor is wrong
             */
            string zUser = Program.entusername;
            string zPassword = Program.entpassword;
            string[] arr1 = new string[] { "", "" };
            string prodDB = extraStuff.getMC_DB();                 // String used to get the server name and the ent db into the array
            arr1 = prodDB.Split('|');
            prodDB = arr1[0];
            string zServer = arr1[1];
            string sql = "";
            SqlConnection conn = new SqlConnection("Data Source=" + zServer + ";Initial Catalog=" + prodDB + "; Integrated Security=False; User ID=" + zUser + "; Password=" + zPassword);
            SqlCommand cmd = new SqlCommand(sql, conn);
            string zResult = "";
            int rowCnt = 0;

            //Connect to the entity database
            try
            {
                conn.Open();
                if (conn.State== ConnectionState.Open)
                {
                   
                }
            }
            catch
            {
                //Do nothing
            }
            cmd.CommandType = CommandType.Text;
            cmd.CommandTimeout = 480;
            cmd.Parameters.Clear();
            switch (ErrCode)
            {
                case 1:
                    /// NOPOPK need to know if because vendor ID does not exit or PO does not exist
                    sql = "select count(*) from PurchaseOrder with (nolock) where POID=" + zPONum + ";";
                    cmd.CommandText = sql;
                    rowCnt = (int)cmd.ExecuteScalar();
                    Console.WriteLine("RowCount is: " + rowCnt.ToString());
                    if (rowCnt == 0)
                    {
                        zResult = "insert into MC_InterfaceLog (ProcessDate, FileName, ErrorMessage, Processed) values (GetDate(), '" + zData + "', '1- PO ID: " + zPONum + " does not exist in the database','N');";
                    }
                    else
                    {
                        zResult = "insert into MC_InterfaceLog (ProcessDate, FileName, ErrorMessage, Processed) values (GetDate(), '" + zData + "', '1- PO ID: " + zPONum + " was found but the Vendor did not match','N');";
                    }
                    break;
                case 2:
                    zResult = "insert into MC_InterfaceLog (ProcessDate, FileName, ErrorMessage, Processed) values (GetDate(), '" + zData + "', '2- Error inserting PO receipt', 'N')";
                    break;
                case 3:
                    zResult = "insert into MC_InterfaceLog (ProcessDate, FileName, ErrorMessage, Processed) values (GetDate(), '" + zData + "', '3- Error getting the PKInvoice', 'N')";
                    break;
                case 4:
                    zResult = "insert into MC_InterfaceLog (ProcessDate, FileName, ErrorMessage, Processed) values (GetDate(), '" + zData + "', '4- Error getting the PK from Invoice or line item is null', 'N')";
                    break;
                default :
                    zResult = "insert into MC_InterfaceLog (ProcessDate, FileName, ErrorMessage, Processed) values (GetDate(), '" + zData + "', '5- Generic error', 'N')";
                    break;           
            }

            // Now let's put that stuff in the database
            cmd.CommandType = CommandType.Text;
            cmd.CommandTimeout = 480;
            cmd.Parameters.Clear();
            cmd.CommandText = zResult;
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch
            {
                //do nothing
            }
            conn.Close();
        }

    }
}
