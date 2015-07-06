using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        static Excel.Application _Excel = null;

        public Form1()
        {
            InitializeComponent();

            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            // Insert code to read the stream here.
                            FileStream fs = myStream as FileStream;
                            if (fs != null)
                            {
                                textBox1.Text = fs.Name.ToString();
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.FileName = "";
            saveFileDialog1.DefaultExt = ".xls";
            saveFileDialog1.Filter = "Excel 文件檔(.xls)|*.xls";

            saveFileDialog1.ShowDialog();

            textBox2.Text = saveFileDialog1.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                //donothing
            }
            else
            {
                this.initailExcel();
                this.openExcel();
            }
        }

        private void initailExcel()
        {
            //檢查PC有無Excel在執行
            bool flag = false;
            foreach (var item in Process.GetProcesses())
            {
                if (item.ProcessName == "EXCEL")
                {
                    flag = true;
                    break;
                }
            }

            if (!flag)
            {
                _Excel = new Excel.Application();
            }
            else
            {
                object obj = Marshal.GetActiveObject("Excel.Application");//引用已在執行的Excel
                _Excel = obj as Excel.Application;
            }

            _Excel.Visible = true;//設false效能會比較好
        }

        void openExcel()
        {
            Excel.Workbook book = null;
            Excel.Range range = null;

            string path = textBox1.Text;
            try
            {
                book = _Excel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//開啟舊檔案
                Excel.Sheets excelSheets = _Excel.Worksheets;
                string currentSheet = "Sheet1";
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets.get_Item(currentSheet);
                //var cellValue = (string)(excelWorksheet.Cells[10, 2] as Excel.Range).Value;

                range = excelWorksheet.UsedRange;
                int lastUsedRow = range.Row + range.Rows.Count - 1;

                for (int r = 1; r < lastUsedRow; ++r)
                {
                    var ProductId = (excelWorksheet.Cells[r, 15] as Excel.Range).Value;
                    try
                    {
                        //數量=2 已處理略過
                        int SaleCount = Convert.ToInt32((string)(excelWorksheet.Cells[r, 20] as Excel.Range).Value);
                        if (SaleCount == 2)
                        {
                            continue;
                        }

                        //2015/05/13~2015/05/15 森永/阿奇儂買一送一
                        //2015/05/22~2015/05/25 森永/阿奇儂買一送一
                        //2015/06/18~2015/07/31 森永/阿奇儂買一送一
                        if (this.CheckEventProduct(Convert.ToInt32(ProductId), 10000277, 10000279) ||
                            this.CheckEventProduct(Convert.ToInt32(ProductId), 10001259, 10001262))
                        {
                            var date = Convert.ToDateTime((excelWorksheet.Cells[r, 25] as Excel.Range).Value);
                            if (this.CheckEventDate(date, new DateTime(2015, 5, 13), new DateTime(2015, 5, 15)) ||
                                this.CheckEventDate(date, new DateTime(2015, 5, 22), new DateTime(2015, 5, 25)) ||
                                this.CheckEventDate(date, new DateTime(2015, 6, 18), new DateTime(2015, 7, 31)))
                            {
                                this.handleEvent(excelWorksheet, r);
                            }
                        }

                        //2015/05/01~2015/05/08 義美買一送一
                        if (this.CheckEventProduct(Convert.ToInt32(ProductId), 10000977, 10000981))
                        {
                            var date = Convert.ToDateTime((excelWorksheet.Cells[r, 25] as Excel.Range).Value);
                            if (this.CheckEventDate(date, new DateTime(2015, 5, 1), new DateTime(2015, 5, 8)))
                            {
                                this.handleEvent(excelWorksheet, r);
                            }
                        }

                        //2015/04/24~2015/04/26 義美買一送一
                        //2015/04/13~2015/04/14 義美買一送一                    10000977, 10000981
                        if (this.CheckEventProduct(Convert.ToInt32(ProductId), 10000968, 10001000) ||
                            this.CheckEventProduct(Convert.ToInt32(ProductId), 10001155, 10001158))
                        {
                            var date = Convert.ToDateTime((excelWorksheet.Cells[r, 25] as Excel.Range).Value);
                            if (this.CheckEventDate(date, new DateTime(2015, 4, 13), new DateTime(2015, 4, 14)) ||
                                this.CheckEventDate(date, new DateTime(2015, 4, 24), new DateTime(2015, 4, 26)))
                            {
                                this.handleEvent(excelWorksheet, r);
                            }
                        }

                    }
                    catch (System.Exception ex)
                    {
                    }

                    //處理momo 新增的贈品欄
                    //var momoPID = (string)(excelWorksheet.Cells[r, 16] as Excel.Range).Value;
                    if (ProductId == "")
                    {
                        string cell = "A" + r.ToString();
                        range = (Excel.Range)excelWorksheet.get_Range(cell, Type.Missing);
                        range.EntireRow.Delete(Excel.XlDirection.xlUp);
                        --r;
                    }

                }

                string savefilename = textBox2.Text;
                book.SaveAs(savefilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            finally
            {
                book.Close(Type.Missing, Type.Missing, Type.Missing);
                book = null;
            }
        }

        void handleEvent(Excel.Worksheet excelWorksheet, int r)
        {
            //變更數量 (買一送一，數量*2)
            excelWorksheet.Cells[r, 20].Value = "2";
            //excelWorksheet.Cells[r, 20].SetValue("2");
            var SellCount = (excelWorksheet.Cells[r, 20] as Excel.Range).Value;

            //變更價格 (買一送一，價錢/2)
            double price = Convert.ToDouble((string)(excelWorksheet.Cells[r, 21] as Excel.Range).Value);
            excelWorksheet.Cells[r, 21].Value = (price / 2).ToString();
            //excelWorksheet.Cells[r, 21].SetValue((price / 2).ToString());
            var SellPrice = (excelWorksheet.Cells[r, 21] as Excel.Range).Value;
        }

        bool CheckEventProduct(int target, int rangelow, int ranghi)
        {
            return target >= rangelow && target <= ranghi;
        }

        bool CheckEventProduct(int target, int PID)
        {
            return target == PID;
        }

        bool CheckEventDate(DateTime OrderDate, DateTime EventDate_start, DateTime EventDate_end)
        {
            return (DateTime.Compare(OrderDate, EventDate_start) >= 0 && DateTime.Compare(OrderDate, EventDate_end) <= 0);
        }
    }
}
