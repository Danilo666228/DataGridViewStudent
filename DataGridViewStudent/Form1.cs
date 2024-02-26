using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
namespace DataGridViewStudent
{
    public partial class MainForm : Form
    {
        Random rand = new Random();
        public MainForm()
        {
            InitializeComponent();
        }
        private void ReadSubjectForFile()
        {
            using (var reader = new StreamReader("Subject.txt"))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    bool isNumber = false;
                    foreach (char c in line)
                    {
                        if (char.IsDigit(c))
                        {
                            isNumber = true;
                            break;
                        }
                    }
                    if (!isNumber)
                    {
                        DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                        column.HeaderText = line;
                        column.Width = 160;
                        DGW.Columns.Add(column);
                    }
                }
            }
        }
        private void ReadFullNameForFile(string path)
        {
            using (var reader = new StreamReader(path))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    bool isNumber = false;
                    foreach (char i in line)
                    {
                        if (char.IsDigit(i))
                        {
                            isNumber = true;
                            break;
                        }
                    }
                    if (!isNumber)
                    {
                        DGW.Rows.Add(line);
                    }
                }
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.SelectedIndex)
            {
                case 0:
                    DGW.Visible = true;
                    ReadFullNameForFile("ИП-21-7К.txt");
                    ReadSubjectForFile();
                    break;
                case 1:
                    DGW.Visible = true;
                    ReadFullNameForFile("ИП-22-1.txt");
                    ReadSubjectForFile();
                    break;
                case 2:
                    DGW.Visible = true;
                    ReadFullNameForFile("СР-21-2.txt");
                    ReadSubjectForFile();
                    break;
            }
        }
        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(comboBox2.Text))
            {
                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                column.HeaderText = comboBox2.Text;
                column.Width = 160;
                DGW.Columns.Add(column);
            }
            else
            {
                MessageBox.Show("Вы не выбрали предмет", "Ошибка",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                DGW.Columns.RemoveAt(DGW.Columns.Count - 1);
            }
            catch
            {
                MessageBox.Show("Колонок больше нет.");
            }
        }
        private void btnFillGrade_Click(object sender, EventArgs e)
        {
            for (int i = 1; i < DGW.ColumnCount; i++)
            {
                for (int j = 0; j < DGW.RowCount - 1; j++)
                {
                    DGW[i, j].Value = rand.Next(2, 6);
                }
            }
        }
        private void btnAverage_Click(object sender, EventArgs e)
        {
            double[] average = new double[DGW.ColumnCount];
            for (int i = 1; i < DGW.ColumnCount; i++)
            {
                double result = 0;
                for (int j = 0; j < DGW.RowCount - 1; j++)
                {
                    result += Convert.ToDouble((DGW[i, j].Value));
                }
                average[i - 1] = result / (DGW.RowCount - 1);
            }
            for (int i = 1; i < DGW.ColumnCount; i++)
            {
                DGW[i, DGW.RowCount - 1].Value = Math.Round(average[i - 1], 2);

            }
        }
        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            var application = new Excel.Application();
            application.SheetsInNewWorkbook = 1;
            Excel.Workbook wb = application.Workbooks.Add(Type.Missing);
            Excel.Worksheet worksheet = application.Worksheets.Item[1];
            worksheet.Name = "Свободная ведомость";
            for (int i = 0; i < DGW.ColumnCount; i++)
            {
                string columnName = DGW.Columns[i].HeaderText;
                worksheet.Cells[1, i + 1] = columnName;
                for (int j = 0; j < DGW.RowCount; j++)
                {
                    worksheet.Cells[j + 2, i + 1] = DGW.Rows[j].Cells[i].Value;
                }
            }
            application.Visible = true;
        }

        private void btnExportWord_Click(object sender, EventArgs e)
        {
            var application = new Word.Application();
            Word.Document document = application.Documents.Add();
            Word.Paragraph userParagraph = document.Paragraphs.Add();
            Word.Range userRange = userParagraph.Range;
            userRange.Text = "Свободная ведомость";
            userRange.InsertParagraphAfter();
            Word.Paragraph tableParagraph = document.Paragraphs.Add();
            Word.Range tableRange = userParagraph.Range;
            Word.Table numbersTable = document.Tables.Add(tableRange, DGW.ColumnCount, 2);
            numbersTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            numbersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            Word.Range cellRange;
            cellRange = numbersTable.Cell(1, 1).Range;
            for (int i = 0; i < DGW.Columns.Count; i++)
            {
                for (int j = 1; j < DGW.Rows.Count; i++)
                {
                    
                    cellRange.Text = DGW[i, j].Value.ToString();
                }

            }


            application.Visible = true;
        }
    }
}
