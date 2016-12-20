using System;
using System.Data;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {

        private string Excel03ConString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private string Excel07ConString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR={1}'";
        private int numRow;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void DisplaySheet()
        {
            string filePath = openFileDialog1.FileName;
            string extension = Path.GetExtension(filePath);
            string header = "NO";
            string conStr, sheetName;

            conStr = string.Empty;
            switch (extension)
            {

                case ".xls": //Excel 97-03
                    conStr = string.Format(Excel03ConString, filePath, header);
                    break;

                case ".xlsx": //Excel 07
                    conStr = string.Format(Excel07ConString, filePath, header);
                    break;
            }

            //Get the name of the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    cmd.Connection = con;
                    con.Open();
                    DataTable dtExcelSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                    con.Close();
                }
            }

            //Read Data from the First Sheet.
            using (OleDbConnection con = new OleDbConnection(conStr))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    using (OleDbDataAdapter oda = new OleDbDataAdapter())
                    {
                        DataTable dt = new DataTable();
                        cmd.CommandText = "SELECT * From [" + sheetName + "]";
                        cmd.Connection = con;
                        con.Open();
                        oda.SelectCommand = cmd;
                        oda.Fill(dt);
                        con.Close();

                        //Populate DataGridView.
                        dataGridView1.DataSource = dt;
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        { 
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Form1.ActiveForm.Text = openFileDialog1.FileName.ToString();
                DisplaySheet();
                numRow = dataGridView1.RowCount;
                label1.Text = "There are " + numRow.ToString() + " records in the Spreadsheet";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string sPath = "c:/temp/PartialFile_Last.sql";
            int curRec = 0;
            int zThou = 0;

            progressBar1.Step = 1;
            progressBar1.Maximum = numRow;
            panel1.Visible= true;
            this.Refresh();
            listBox1.Items.Add("DECLARE                                                            ");
            listBox1.Items.Add("     @RC_PK   VARCHAR(20)                                          ");
            listBox1.Items.Add("    ,@RC_Name VARCHAR(50)                                          ");
            listBox1.Items.Add("    ,@RC_ID   VARCHAR(10)                                          ");
            listBox1.Items.Add("    ,@AssPK INT                                                    ");
            listBox1.Items.Add("    ,@NextPK  INT;                                                 ");
            listBox1.Items.Add("                                                                   ");
            foreach (DataGridViewRow row in dataGridView1.Rows) 
            {
                try
                {                    //zSQL = "select * from document where AssetPK= (select AssetPK from Asset where AssetID = '"+ row.Cells[4].Value.ToString() + "')";
                    listBox1.Items.Add("-------------------------- NEW Doc --------------------------------");
                    listBox1.Items.Add("Set @AssPK= '"+ row.Cells[4].Value.ToString() + "'                 ");
                    listBox1.Items.Add("                                                                   ");
                    listBox1.Items.Add("SELECT @RC_PK = RepairCenterPK                                     ");
                    listBox1.Items.Add("      ,@RC_Name = RepairCenterName                                 ");
                    listBox1.Items.Add("      ,@RC_ID = RepairCenterID                                     ");
                    listBox1.Items.Add("FROM                                                               ");
                    listBox1.Items.Add("     Asset a                                                       ");
                    listBox1.Items.Add("WHERE  AssetPK = @AssPK;                                           ");
                    listBox1.Items.Add("                                                                   ");
                    listBox1.Items.Add("INSERT INTO Document WITH(ROWLOCK)                                 ");
                    listBox1.Items.Add("  (DocumentID                                                      ");
                    listBox1.Items.Add("  ,DocumentName                                                    ");
                    listBox1.Items.Add("  ,RepairCenterPK                                                  ");
                    listBox1.Items.Add("  ,RepairCenterID                                                  ");
                    listBox1.Items.Add("  ,RepairCenterName                                                ");
                    listBox1.Items.Add("  ,DocumentType                                                    ");
                    listBox1.Items.Add("  ,DocumentTypeDesc                                                ");
                    listBox1.Items.Add("  ,LocationType                                                    ");
                    listBox1.Items.Add("  ,LocationTypeDesc                                                ");
                    listBox1.Items.Add("  ,Location                                                        ");
                    listBox1.Items.Add("  ,DisplayLink                                                     ");
                    listBox1.Items.Add("  ,PrintWithWO                                                     ");
                    listBox1.Items.Add("  ,SentWithEmail                                                   ");
                    listBox1.Items.Add("  ,Active                                                          ");
                    listBox1.Items.Add("  ,RowVersionDate                                                  ");
                    listBox1.Items.Add("  ,RowVersionInitials                                              ");
                    listBox1.Items.Add("  )                                                                ");
                    listBox1.Items.Add("VALUES                                                             ");
                    listBox1.Items.Add("  ('toBrplced'                                                     ");
                    listBox1.Items.Add("  ,'"+ row.Cells[2].Value.ToString()  + "'                         ");
                    listBox1.Items.Add("  ,@RC_PK                                                          ");
                    listBox1.Items.Add("  ,@RC_ID                                                          ");
                    listBox1.Items.Add("  ,@RC_Name                                                        ");
                    listBox1.Items.Add("  ,'INFO'                                                          ");
                    listBox1.Items.Add("  ,'Informational Document'                                        ");
                    listBox1.Items.Add("  ,'HTTPLIBRARY'                                                   ");
                    listBox1.Items.Add("  ,'Library Link'                                                  ");
                    listBox1.Items.Add("  ,'" + row.Cells[0].Value.ToString() + "'                         ");
                    listBox1.Items.Add("  ,0                                                               ");
                    listBox1.Items.Add("  ,0                                                               ");
                    listBox1.Items.Add("  ,0                                                               ");
                    listBox1.Items.Add("  ,1                                                               ");
                    listBox1.Items.Add("  ,GETDATE()                                                       ");
                    listBox1.Items.Add("  ,'_MC'                                                           ");
                    listBox1.Items.Add("  )                                                                ");
                    listBox1.Items.Add("                                                                   ");
                    listBox1.Items.Add("SET @NextPK = SCOPE_IDENTITY()                                     ");
                    listBox1.Items.Add("                                                                   ");
                    listBox1.Items.Add("INSERT INTO AssetDocument                                          ");
                    listBox1.Items.Add("  (DocumentPK                                                      ");
                    listBox1.Items.Add("  ,AssetPK                                                         ");
                    listBox1.Items.Add("  ,ModuleID                                                        ");
                    listBox1.Items.Add("  ,DisplayLink                                                     ");
                    listBox1.Items.Add("  ,PrintWithWO                                                     ");
                    listBox1.Items.Add("  ,SendWithEmail                                                   ");
                    listBox1.Items.Add("  )                                                                ");
                    listBox1.Items.Add("VALUES                                                             ");
                    listBox1.Items.Add("  (@NextPK                                                         ");
                    listBox1.Items.Add("  ,@AssPK                                                          ");
                    listBox1.Items.Add("  ,'AS'                                                            ");
                    listBox1.Items.Add("  ,0                                                               ");
                    listBox1.Items.Add("  ,0                                                               ");
                    listBox1.Items.Add("  ,0                                                               ");
                    listBox1.Items.Add("  )                                                                ");
                    listBox1.Items.Add("                                                                   ");
                    listBox1.Items.Add("UPDATE Document                                                    ");
                    listBox1.Items.Add("SET                                                                ");
                    listBox1.Items.Add("   DocumentID = DocumentPK                                         ");
                    listBox1.Items.Add("WHERE                                                              ");
                    listBox1.Items.Add("       DocumentPK = @NextPK                                        ");
                    progressBar1.PerformStep();
                    curRec++;
                    if (curRec % 1000 == 0)
                    {
                        zThou++;
                        label4.Text = zThou.ToString();
                        panel1.Refresh();
                        StreamWriter SavePartialFile = new StreamWriter("c:/Temp/PartialFile_"+ zThou.ToString() +".sql");
                        foreach (var item in listBox1.Items)
                        {
                            SavePartialFile.WriteLine(item.ToString());
                        }
                        SavePartialFile.ToString();
                        SavePartialFile.Close();
                        SavePartialFile.Dispose();
                        listBox1.Items.Clear();
                    } 
                }
                catch { };
            }


            StreamWriter SaveFile = new StreamWriter(sPath);
            foreach (var item in listBox1.Items)
            {
                SaveFile.WriteLine(item.ToString());
            }
            SaveFile.ToString();
            SaveFile.Close();
            SaveFile.Dispose();
            panel1.Visible = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panel1.Visible = false;
        }
    }
}
