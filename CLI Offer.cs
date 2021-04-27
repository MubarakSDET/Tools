using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CLI_Offer
{
    public partial class Form1 : Form
    {
        public static string CIF = null;
        public static string AccNo = null;

        public static string Env = null;
        public static string Selectedfile = null;

       

        Program myprgm = new Program();
        public Form1()
        {
            InitializeComponent();
        }

       

        public void myprgm_Progress(int value, int total)
        {
            //progressBar1.Value = value;
            toolStripProgressBar1.Value = value;
            toolStripProgressBar1.Maximum = total;
        }

        public void myprgm_SetToolStatus(string logtxt)
        {
            //progressBar1.Value = value;
            toolStripStatusLabel1.Text = logtxt;
            //toolStripProgressBar1.Value = value;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            linkLabel1.Enabled = false;

            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.label2, "For Tool Support Contact: DEADODataServices&BatchOperations@cba.com.au");
            
        }
        private void button3_Click(object sender, EventArgs e)
        {
            string Today = DateTime.Today.ToString();
            string Time = DateTime.Now.ToString("HH:mm:ss tt");
            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string Toolname = "CLI Offer";
            string Appln = "NETBANK";  

            if (!string.IsNullOrEmpty(Selectedfile))
            {

                if ((radioButton1.Checked == true) || (radioButton2.Checked == true))
                {

                    string result1 = null;
                    foreach (Control control in this.groupBox1.Controls)
                    {
                        if (control is RadioButton)
                        {
                            RadioButton radio = control as RadioButton;
                            if (radio.Checked) { result1 = radio.Text; }
                        }
                    }

                    
                        myprgm_SetToolStatus("Executing ...");
                   

                    if (radioButton1.Checked == true)
                    {
                        Env = "T2";
                    }

                    if (radioButton2.Checked == true)
                    {
                        Env = "T5";
                    }

                    
                        myprgm.ReadExcelArgument(tbFilePath.Text, result1);
                    

                    linkLabel1.Enabled = true;
                    System.Threading.Thread.Sleep(300);

                    MessageBox.Show("Executed! The Results have been updated in the Excel.\nPlease click the Exit button in the application to exit!", "You are Done!",
                    MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    myprgm_SetToolStatus("Execution completed");

                    myprgm.LogResults(myprgm.curtime, userName, Appln, Env, Toolname, myprgm.noofTransactions);
                }
                else
                {
                    MessageBox.Show("Please select the Environment.");
                    
                }

            }
            else
            {
                MessageBox.Show("Please select the Data file.");
            }
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(Selectedfile);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {

                tbFilePath.Text = openFileDialog1.FileName;
                Selectedfile = tbFilePath.Text;
                myprgm_SetToolStatus("Input File Selected");

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            tbFilePath.Clear();
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            linkLabel1.Enabled = false;
            myprgm_SetToolStatus("Online");
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

       


    }
}
