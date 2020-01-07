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
            List<List<string>> tmp = new List<List<string>>();
            ObjWorkSheet = (Excel.Worksheet)workBook.Sheets[1];
            tmp = new List<List<string>>();
            Excel.Range cells = ObjWorkSheet.Cells;
            var lastCell = cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            for (int i = 0; i < lastCell.Column; i++)
            {
                tmp.Add(new List<string>());
                for (int j = 1; j < lastCell.Row; j++)
                    tmp[i].Add(cells[j + 1, i + 1].Text);
            }
            workBook.Close(true);
            exApp.Quit();
            int kolvoStudGrup;
            int kolStud = tmp[0].Count;
           int kolvoGrup = Convert.ToInt32(Convert.ToInt32(textBox1.Text));
            kolvoStudGrup = kolStud / kolvoGrup;
            int x=0, y=0;
            for (int i = 0; i < kolvoGrup; i++)
            {
                if (i > 0)
                {
                    x += 450;
                    y = 0;
                }
                for (int j = 0; j < kolvoStudGrup; j++)
                {
                Students Stud = new Students
                {
                    Location = new Point(15 + x, 15 + y)
                };
                button1.Visible = false;
                button2.Visible = false;
                label1.Visible = true;
                textBox1.Visible = true;
                string[] mas = tmp[0][i].Split(' ');
                Stud.textBox1.Text = mas[0];
                Stud.textBox2.Text = mas[1];
                Stud.label4.Text = j.ToString();
                if (mas.Length == 3)
                    Stud.textBox3.Text = mas[2];
                
                Controls.Add(Stud);
                    y += 60;
                }

                

            }
        }
    }
}
           