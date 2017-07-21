using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Configuration;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FireSomething
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void loadZPLFile(string zFileName)
        {
            Process myProcess = new Process();

            try
            {
                myProcess.StartInfo.UseShellExecute = false;
                myProcess.StartInfo.FileName = textBox2.Text;
                myProcess.StartInfo.Arguments = zFileName;
                myProcess.StartInfo.CreateNoWindow = false;
                myProcess.Start();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
                button3.Enabled= true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog()== DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
                button4.Enabled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            if (textBox1.Text !="")
            {
                DirectoryInfo zDir = new DirectoryInfo(textBox1.Text);
                foreach (FileInfo zFile in zDir.GetFiles("*.zpl"))
                {
                    loadZPLFile(Path.Combine(textBox1.Text, zFile.Name));
                    listView1.View = View.Details;
                    ListViewItem newItem = new ListViewItem(DateTime.Now.ToString());
                    newItem.SubItems.Add(zFile.Name);
                    listView1.Items.Add(newItem); 
                    label3.Text = "There are " + listView1.Items.Count.ToString() + " ZPL files";
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = ConfigurationSettings.AppSettings["FolderProcess"];      
            textBox2.Text = ConfigurationSettings.AppSettings["ProcessExecutable"];
        }
    }
}
