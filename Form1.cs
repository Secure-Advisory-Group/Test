using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Proftaak_Software
{
    public partial class Form1 : Form
    {
        string KlantenDatabase = "{0, -15}{1, -30}{2, -45}";
        int[] Kolom = new int[0];
        int[] Rij = new int[0];
        public Form1()
        {
            InitializeComponent();
        }

        public void OpenFile()
        {
            Excel excel = new Excel(Convert.ToString(listBox2.SelectedItem), 1);
            
            string Bedrijfsnaam, Locatie, Veiligheidsrisico;
            for (int i = 0; i < excel.ColumnCellAmount(); i++)
            {
                Bedrijfsnaam = excel.ReadCell(i, 0);
                Locatie = excel.ReadCell(i, 1);
                Veiligheidsrisico = excel.ReadCell(i, 2);
                listBox1.Items.Add(string.Format(KlantenDatabase, Bedrijfsnaam, Locatie, Veiligheidsrisico));
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            OpenFile();
        }

        public void WriteData()
        {
            Excel excel = new Excel(@"C:\Users\Moshe\Documents\TEST1", 1);

            excel.WriteToCell(0, 0, "Test2");
            excel.Save();
            excel.SaveAs(@"C:\Users\Moshe\Documents\Test2");

            excel.Close();
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            WriteData();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listBox1.Items.Add(string.Format(KlantenDatabase, "Bedrijfsnaam", "Locatie", "Veiligheidsrisico"));
        }
        
        private void Button3_Click_1(object sender, EventArgs e)
        {
            {
                FolderBrowserDialog FBD = new FolderBrowserDialog();

                if (FBD.ShowDialog() == DialogResult.OK)
                {
                    listBox2.Items.Clear();
                    string[] files = Directory.GetFiles(FBD.SelectedPath);
                    string[] dirs = Directory.GetDirectories(FBD.SelectedPath);

                    foreach (string file in files)
                    {
                        listBox2.Items.Add(file);
                    }
                    foreach (string dir in dirs)
                    {
                        listBox2.Items.Add(dir);
                    }
                }
            }
        }
    }
}
