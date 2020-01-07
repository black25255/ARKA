using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApp5
{
    public partial class Form1 : Form
    {
        
        Excel.Application exApp = new Excel.Application();

        public Form1()
        {
            InitializeComponent();
        }

        void excel()
        {
            Excel.Application ExApp = new Excel.Application();
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    Form r = new BarSliyanie();
                    r.Show();
                    //(Application.OpenForms["BarSliyanie"] as BarSliyanie).info("Чтение файла", 10);

                    List<string> fam = new List<string>();
                    List<string> nam = new List<string>();
                    List<string> otch = new List<string>();
                    ofd.Title = "Открыть базу ...";
                    ofd.Filter = "*.xlsx|*.xlsx;|*.xls|*.xls;|*.xml|";
                    if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        Text = ofd.FileName;
                    }
                    Excel.Workbook workBook = exApp.Workbooks.Open(ofd.FileName);
                    Excel.Worksheet ObjWorkSheet;
                    List<List<string>>[] tmp = new List<List<string>>[6];
                    for (int k = 0; k < 6; k++)
                    {
                        ObjWorkSheet = (Excel.Worksheet)workBook.Sheets[k + 1];
                        tmp[k] = new List<List<string>>();
                        var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                        for (int i = 0; i < lastCell.Row; i++)
                        {
                            tmp[k].Add(new List<string>());
                            for (int j = 0; j < lastCell.Row; j++)
                                tmp[k][i].Add(ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString());
                        }
                    }
                    workBook.Close(true);
                    exApp.Quit();


                }
            }
            catch
            {
                MessageBox.Show("Проблема с файлом");
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int i = Convert.ToInt32(((Button)(sender)).Tag);


            OpenFileDialog sfd = new OpenFileDialog();

            sfd.Title = "Открыть базу ....";
            sfd.Filter = "*.xlsx|*.xlsx;|*.xls|*.xls|*.xml|";
            sfd.AddExtension = true;

            if (sfd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Text = sfd.FileName;

            }
            Excel.Workbook workBook = exApp.Workbooks.Open(sfd.FileName);
            Excel.Worksheet ObjWorkSheet;
            List<List<string>>[] tmp = new List<List<string>>[1];
            for (int k = 0; k < 1; k++)
            {
                ObjWorkSheet = (Excel.Worksheet)workBook.Sheets[k + 1];
                tmp[k] = new List<List<string>>();
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                for (int l = 0; l < lastCell.Column;l++)
                {
                    tmp[k].Add(new List<string>());
                        
                }

            }

            

        }

        private void students1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<List<string>>[] tmp = new List<List<string>>[1];
            students students1 = new students();
            students1.Location = new Point(12, 60);
            this.Controls.Add(students1);

            students1.textBox1.Text = "ФИО";
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }
    }
}