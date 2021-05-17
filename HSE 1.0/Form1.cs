using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace HSE_1._0
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        public void sendMessage(string msg) {

            MessageBox.Show(msg);
        }

        private void Close_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void Open_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Multiselect = true;

            DialogResult dr = openFileDialog1.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                string[] locationArray = openFileDialog1.FileNames;

                HSEReport passFilePath = new HSEReport();
                string value = textBox1.Text;
                passFilePath.openFile(locationArray, value);
            }

            // another option to open dialog

            /*openFileDialog1.Multiselect = false;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string fileName = Path.GetFileName(openFileDialog1.FileName);
                string filePath = openFileDialog1.FileName;
                string[] locationArray = openFileDialog1.FileNames;
                HSEReport passFilePath = new HSEReport();
                //passFilePath.openFile(locationArray);
                MessageBox.Show(locationArray[0]);
            }*/

            progressBar1.Value = 0;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Step = 1;

            for (int i = 0; i < 100; i++)
            {
                progressBar1.PerformStep();
            }
        }

        public void textBox1_TextChanged(object sender, EventArgs e)
        {
            string value = textBox1.Text;
            /*MessageBox.Show(value.ToString());*/
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
