using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ASPackUnpacker
{
    public partial class mainFrm : Form
    {
        public mainFrm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Executable files only (*.exe)|*.exe";
            if (open.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = open.FileName;
            }
        }

        public void AddLog(string toAdd)
        {
            listBox1.Items.Add(DateTime.UtcNow.ToShortTimeString() + " ->" + toAdd);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 0)
            {
                if (File.Exists(textBox1.Text))
                {
                    AddLog("Starting unpacking...");
                    Unpacker unp = new Unpacker(textBox1.Text);
                    unp.Unpack(this);
                }
                else
                {
                    AddLog("File doesn't exist!");
                }
            }
            else
            {
                AddLog("Please select a file first!");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void mainFrm_Load(object sender, EventArgs e)
        {
            AddLog("NIDebugger::ASPackUnpacker loaded...");
        }
    }
}
