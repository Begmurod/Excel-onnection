using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelMacrosExample
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

       

        private void button_Calc_Click_1(object sender, EventArgs e)
        {
            Excel.Application objExcel = null;
            Excel.Workbook WorkBook = null;

            try
            {
                objExcel = new Excel.Application();

                //если надо показать Excel-файл
                //objExcel.ScreenUpdating = true;
                //objExcel.WindowState = Excel.XlWindowState.xlMaximized;
                //objExcel.Visible = true;
                //objExcel.DisplayAlerts = true;

                string fileName = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "Test.xlsm");

                WorkBook = objExcel.Workbooks.Open(fileName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);

                Excel.Worksheet WorkSheet = (Excel.Worksheet)WorkBook.Sheets["Test"];
                WorkSheet.Range["B5"].Value2 = Convert.ToDouble(textBox_lamOGH.Text);
                WorkSheet.Range["B6"].Value2 = Convert.ToDouble(textBox_lam1.Text);
                WorkSheet.Range["B7"].Value2 = Convert.ToDouble(textBox_lam2.Text);
                WorkSheet.Range["B9"].Value2 = Convert.ToDouble(textBox_c1.Text);
                WorkSheet.Range["B10"].Value2 = Convert.ToDouble(textBox_c2.Text);
                WorkSheet.Range["B11"].Value2 = Convert.ToDouble(textBox_Mct.Text);
                WorkSheet.Range["B13"].Value2 = Convert.ToDouble(textBox_t11.Text);
                WorkSheet.Range["B14"].Value2 = Convert.ToDouble(textBox_t22.Text);
                WorkSheet.Range["B15"].Value2 = Convert.ToDouble(textBox_tmax.Text);
                WorkSheet.Range["B17"].Value2 = Convert.ToDouble(textBox_alfa1.Text);
                WorkSheet.Range["B18"].Value2 = Convert.ToDouble(textBox_alfa2.Text);
                WorkSheet.Range["B19"].Value2 = Convert.ToDouble(textBox_Cz.Text);
                WorkSheet.Range["B20"].Value2 = Convert.ToDouble(textBox_To.Text);
                WorkSheet.Range["B21"].Value2 = Convert.ToDouble(textBox_Tol.Text);
                WorkSheet.Range["B22"].Value2 = Convert.ToDouble(textBox_To2.Text);
                WorkSheet.Range["B23"].Value2 = Convert.ToDouble(textBox_To1.Text);
                WorkSheet.Range["B24"].Value2 = Convert.ToDouble(textBox_To.Text);


                objExcel.GetType().InvokeMember("Run", BindingFlags.Default | BindingFlags.InvokeMethod,
                    null, objExcel, new Object[] { "Test" });
                textBox_Cz.Text = WorkSheet.Range["B19"].Value.ToString("0.##");
                textBox_Summ.Text = WorkSheet.Range["B20"].Value.ToString("0.##");
                textBox_To.Text = WorkSheet.Range["B24"].Value.ToString("0.##");
                textBox_To2.Text = WorkSheet.Range["B22"].Value.ToString("0.##");
                textBox_To1.Text = WorkSheet.Range["B23"].Value.ToString("0.##");
                

                MessageBox.Show("Решение найдено.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка вызова (" + ex.Message + ").");
            }
            finally
            {
                if (WorkBook != null) WorkBook.Close(false, null, null);
                if (objExcel != null) objExcel.Quit();
            }

        }
    }
}
