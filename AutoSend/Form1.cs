using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinForms = System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Net.Mail;
using System.Net.Mime;

namespace AutoSend
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        static string pathXls = "";
        static string pathFolder = "";

        public void FileReader()
        {
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(@"" + pathXls);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            int nInLastRow = ObjWorkSheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            int nInLastCol = ObjWorkSheet.Cells.Find("*", System.Reflection.Missing.Value,
                System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;

            string[,] list = new string[nInLastRow, nInLastCol];

            for (int i = 0; i < nInLastRow; i++)
            {
                string[] tempArrStr = new string[nInLastCol];
                for (int j = 0; j < nInLastCol; j++)
                {
                    string str = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString().Replace("\n", "\r\n");
                    list[i, j] = str;
                    tempArrStr[j] += list[i, j];

                    if (j == nInLastCol - 1)
                    {
                        dataGridView1.Rows.Add(tempArrStr);
                    }
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            int k = 0;
            int kk = 0;
            try
            {
                if (MessageBox.Show("Отправить письма уверен ли ты, что хочешь?", "Точна?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        Outlook.Application oApp = new Outlook.Application();
                        Outlook.MailItem mail = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

                        k++;
                        kk = 0;
                        for (;;)
                        {
                            kk = 1;
                            mail.To = dataGridView1.Rows[i].Cells[0].Value.ToString();
                            kk = 3;
                            mail.Subject = dataGridView1.Rows[i].Cells[2].Value.ToString();
                            kk = 4;
                            mail.Body = dataGridView1.Rows[i].Cells[3].Value.ToString();
                            kk = 2;
                            string str = dataGridView1.Rows[i].Cells[1].Value.ToString();
                            string[] attachmentPaths = explode(",", str);
                            for (int j = 0; j < attachmentPaths.Length; j++)
                            {
                                attachmentPaths[j] = @"" + pathFolder + "\\" + attachmentPaths[j] + ".pdf";
                                mail.Attachments.Add(attachmentPaths[j], null, null, null);
                            }
                            mail.Importance = Outlook.OlImportance.olImportanceNormal;
                            mail.Send();
                            oApp = null;
                            mail = null;
                            break;
                        }
                    }
                    MessageBox.Show("Гатова", "Ok", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch
            {
                MessageBox.Show("Чо-то не то с " + k + " строкой, " + kk + " столбцом", "Ой", MessageBoxButtons.OK);
            }
        }

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    textBox1.Text = row.Cells[3].Value.ToString();
                }
            }
            catch
            {

            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            try
            {
                foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                {
                    if (textBox1.Text != null)
                    {
                        cell.Value = textBox1.Text;
                    }
                }
            }
            catch
            {

            }
        }

        private void DataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                {
                    textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                }
            }
            catch
            {

            }
        }

        private string[] explode(string separator, string source)
        {
            return source.Split(new string[] { separator }, StringSplitOptions.None);
        }

        private void Panel1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "Файлы xlsx|*.xlsx";
            ofd.ShowDialog();

            pathXls = ofd.FileName;
            if (pathXls != "" && pathFolder != "")
            {
                FileReader();
                panel1.Visible = false;
                panel2.Visible = false;
                dataGridView1.Visible = true;
                textBox1.Visible = true;
                button1.Visible = true;
                button2.Visible = true;
            }
        }

        private void Panel2_Click(object sender, EventArgs e)
        {
            WinForms.FolderBrowserDialog fbd = new WinForms.FolderBrowserDialog();
            fbd.ShowDialog();

            pathFolder = fbd.SelectedPath;
            if (pathXls != "" && pathFolder != "")
            {
                FileReader();
                panel1.Visible = false;
                panel2.Visible = false;
                dataGridView1.Visible = true;
                textBox1.Visible = true;
                button1.Visible = true;
                button2.Visible = true;
            }
        }

        private void Panel1_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    string file = (string)files[0];

                    pathXls = file;
                }

                if (pathXls != "" && pathFolder != "")
                {
                    FileReader();
                    panel1.Visible = false;
                    panel2.Visible = false;
                    dataGridView1.Visible = true;
                    textBox1.Visible = true;
                    button1.Visible = true;
                    button2.Visible = true;
                }
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Panel2_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    string file = (string)files[0];

                    pathFolder = file;
                }

                if (pathXls != "" && pathFolder != "")
                {
                    FileReader();
                    panel1.Visible = false;
                    panel2.Visible = false;
                    dataGridView1.Visible = true;
                    textBox1.Visible = true;
                    button1.Visible = true;
                    button2.Visible = true;
                }
            }
            catch
            {
                MessageBox.Show("Что-то пошло не так", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Panel1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void Panel2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }
    }
}
