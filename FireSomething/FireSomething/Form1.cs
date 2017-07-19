using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
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
                    listView1.View = View.Details;
                    ListViewItem newItem = new ListViewItem(DateTime.Now.ToString());
                    newItem.SubItems.Add(zFile.Name);
                    listView1.Items.Add(newItem); 
                    label3.Text = "There are " + listView1.Items.Count.ToString() + " ZPL files";
                }
            }
        }
    }
}
