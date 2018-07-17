using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2
{
    public partial class Form1 : Form
    {
        Excel.Application excelApp = null;
        public Excel.Workbook wb;                  // Mappe
        public Excel.Worksheet ws;                 // Tab
        public int GetSet_Columns { get; set; } = 10;

        public Form1()
        {
            InitializeComponent();
        }

        public bool Open()
        {
            try
            {
                excelApp = new Excel.Application();
                excelApp.Visible = true;
                excelApp.ScreenUpdating = true;
                try
                {
                    string path = "";
                    OpenFileDialog Import = new OpenFileDialog();
                    Import.Filter = "Excel-Arbeitsmappe (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    if (Import.ShowDialog() == DialogResult.OK)
                    {
                        path = Import.FileName;
                    }
                    wb = excelApp.Workbooks.Open(path);
                    return true;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Haben Sie Excel installiert?");
                return false;
            }
        }

        int counter = 1;
        private void btnStart_Click(object sender, EventArgs e)
        {
            if (GetSet_Columns > 0)
            {
                int i = 0;
                while (i < GetSet_Columns)
                {
                    ws = (Excel.Worksheet)wb.ActiveSheet;
                    ((Excel.Worksheet)excelApp.ActiveWorkbook.Sheets[1]).Select();
                    Excel.Range oRng = ws.Range[ws.Cells[counter, counter], ws.Cells[counter + 10, counter + 10]];

                    Excel.Range _sortBy = ws.get_Range(tB_Column_From.Text + counter.ToString(), tB_Column_To.Text + counter.ToString());
                    Excel.Range _sortRange = ws.get_Range(tB_Column_From.Text + counter.ToString(), tB_Column_To.Text + counter.ToString());



                    ws.Sort.SortFields.Clear();
                    ws.Sort.SetRange(_sortRange);
                    ws.Sort.SortFields.Add(_sortBy, 0, SortOrder.Ascending);
                    ws.Sort.Header = Excel.XlYesNoGuess.xlYes;
                    ws.Sort.MatchCase = false;
                    ws.Sort.Orientation = Excel.XlSortOrientation.xlSortRows;
                    ws.Sort.SortMethod = Excel.XlSortMethod.xlStroke;
                    ws.Sort.Apply();
                    counter++;
                    i++;
                }
            }
            else
            {
                MessageBox.Show("Please recheck your input on columns, mate.", "Ups!");
            }
            counter = 1;
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            Open();
        }

        private void tB_Columns_TextChanged(object sender, EventArgs e)
        {
            int temp = 0;
            try
            {
                temp = Convert.ToInt32(tB_Columns.Text);
            }
            catch(Exception ex)
            {
                MessageBox.Show("Please only use natural(integer) numbers!", "Error!");
            }
            GetSet_Columns = temp;
        }

        private void tB_Column_From_TextChanged(object sender, EventArgs e)
        {
            tB_Column_From.Text = tB_Column_From.Text.ToUpper();
        }

        private void tB_Column_To_TextChanged(object sender, EventArgs e)
        {
            tB_Column_To.Text = tB_Column_To.Text.ToUpper();
        }
    }
    }
