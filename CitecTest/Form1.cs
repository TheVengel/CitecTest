using System;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace CitecTest
{
    public partial class Form1 : Form
    {
        OpenFileDialog openFileDialog = new OpenFileDialog();
        SaveFileDialog saveFileDialog = new SaveFileDialog();

        private Service service = new Service();

        private string FileNameRRK;
        private string FileNameAppeal;
        private List<Administrant> list;
        private DateTime date = new DateTime();
        private DataTable dt = new DataTable();

        public Form1()
        {
            InitializeComponent();

            dataGridView1.ColumnHeaderMouseClick += dataGridView1_ColumnHeaderMouseClick;
            dataGridView1.DataBindingComplete += dataGridView1_DataBindingComplete;

            openFileDialog.Filter = "Text files(*.txt)|*.txt|All files(*.*)|*.*";
            saveFileDialog.Filter = "Word File (.docx ,.doc)|*.docx;*.doc";

            this.date = DateTime.Today;
            DataTableInit();
        }

        private void DataTableInit()
        {
            DataColumn column;

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "� �.�.";
            column.Caption = "id";



            this.dt.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "������������� �����������";
            column.Caption = "name";

            this.dt.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "���������� ������������� �������� ����������";
            column.Caption = "RRK";

            this.dt.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "���������� ������������� ���������� ��������� �������";
            column.Caption = "Appeal";

            this.dt.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "����� ���������� ���������� � ���������";
            column.Caption = "allDocs";

            this.dt.Columns.Add(column);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.Cancel)
                return;
            
            this.FileNameRRK = openFileDialog.FileName;
            
            label1.Text = this.FileNameRRK;

            this.FillTable();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.Cancel)
                return;

            this.FileNameAppeal = openFileDialog.FileName;

            label2.Text = this.FileNameAppeal;

            this.FillTable();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            
        }
        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            var view = new DataView(dt);
            if (e.ColumnIndex == 1)
                view.Sort = "������������� ����������� desc";
            if (e.ColumnIndex == 2)
                view.Sort = "���������� ������������� �������� ���������� desc, ���������� ������������� ���������� ��������� ������� desc";
            else if (e.ColumnIndex == 3)
                view.Sort = "���������� ������������� ���������� ��������� ������� desc, ���������� ������������� �������� ���������� desc";
            else if (e.ColumnIndex == 4)
                view.Sort = "����� ���������� ���������� � ��������� desc, ���������� ������������� �������� ���������� desc";

            var dtSorted = view;

            dataGridView1.DataSource = view;

            for (var i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                dataGridView1[0, i].Value = i + 1;
            }
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            
        }

        private void FillTable()
        {
            if (FileNameRRK != null && FileNameAppeal != null)
            {
                var stopWatch = new Stopwatch();
                stopWatch.Start();
                service.ReadFile(FileNameRRK, RRK: 1);
                service.ReadFile(FileNameAppeal, Appeal: 1);
                this.list = service.GetList();
                foreach (var item in list)
                    this.dt.Rows.Add(list.IndexOf(item) + 1, item.Name, item.DocsRRK, item.DocsAppeal, item.DocsRRK + item.DocsAppeal);
                stopWatch.Stop();
                dataGridView1.DataSource = this.dt;
                TimeSpan ts = stopWatch.Elapsed;
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:000}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds);
                label4.Text = $"����� ���������� ���������� {elapsedTime}";
                label5.Text = "��� ���������� ������� ���������� ������ �� ��������� �������";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var column = dataGridView1.SortedColumn;
            var sort = "���������� �� �����������"; ;
            if (column != null)
            {
                
                switch (column.Name)
                {
                    case "������������� �����������":
                        {
                            sort = "���������� �� ����� �������������� �����������";
                            break;
                        }
                    case "���������� ������������� �������� ����������":
                        {
                            sort = "���������� �� ���������� ������������� �������� ����������";
                            break;
                        }
                    case "���������� ������������� ���������� ��������� �������":
                        {
                            sort = "���������� �� ���������� ������������� ���������� ��������� �������";
                            break;
                        }
                    case "����� ���������� ���������� � ���������":
                        {
                            sort = "���������� �� ������ ���������� ����������.";
                            break;
                        }
                    default:
                        {
                            break;
                        }
                }
            }

            var wordApp = new Word.Application();
            Word.Document wordDoc = wordApp.Documents.Add();

            Word.Paragraph paragraphTable = wordDoc.Paragraphs.Add();

            Word.Range r0 = paragraphTable.Range;
            r0.Text = $"������� � ������������� ���������� � ���������� ������� \n";
            r0.Font.Bold = 1;
            r0.Font.Size = 16;
            r0.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            Word.Range r1 = paragraphTable.Range;
            r1.Text = $"�� ��������� � ���� {service.GetSumDocs()} ����������, �� ���: \n";
            Word.Range r2 = paragraphTable.Range;
            r2.Text = $"- ���������� ������������� �������� ����������: {service.GetAllRRK()}; \n";
            Word.Range r3 = paragraphTable.Range;
            r3.Text = $"- ���������� ������������� ���������� ��������� �������: {service.GetAllAppeal()}. \n";
            Word.Range r4 = paragraphTable.Range;
            r4.Text = sort + "\n";

            

            Word.Paragraph paragraphTable1 = wordDoc.Paragraphs.Add();

            Word.Range rangeTable = paragraphTable1.Range;
            Word.Table table = wordDoc.Tables.Add(rangeTable, dataGridView1.RowCount, dataGridView1.ColumnCount);
            table.Borders.InsideLineStyle = table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;

            Word.Range cellRange;
            cellRange = table.Cell(1, 1).Range;
            cellRange.Text = "� �.�.";
            cellRange = table.Cell(1, 2).Range;
            cellRange.Text = "������������� �����������";
            cellRange = table.Cell(1, 3).Range;
            cellRange.Text = "���������� ������������� �������� ����������";
            cellRange = table.Cell(1, 4).Range;
            cellRange.Text = "���������� ������������� ���������� ��������� �������";
            cellRange = table.Cell(1, 5).Range;
            cellRange.Text = "����� ���������� ���������� � ���������";
            table.Rows[1].Range.Bold = 1;

            for (var i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                for (var j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    cellRange = table.Cell(i + 2, j + 1).Range;
                    var value = dataGridView1[j, i].Value.ToString();
                    cellRange.Text = value;
                }
            }

            Word.Range dateR = paragraphTable.Range;
            dateR.Text = $"���� ����������� �������:                          {this.date.ToString("d")}";

            if (saveFileDialog.ShowDialog() == DialogResult.Cancel)
                return;

            try
            {
                wordDoc.SaveAs2(saveFileDialog.FileName);
            }
            catch
            {
                MessageBox.Show("�������� ����");
                return;
            }
            wordDoc.Close();
            label3.Text = $"���� �������� �� ����: {saveFileDialog.FileName}";
        }
    }
}