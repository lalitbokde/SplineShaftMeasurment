using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CCD_Measurment_of_SSSA
{
    public partial class Frm_DatewiseReport : Form
    {
        BAL _bal = new BAL();

        public Frm_DatewiseReport()
        {
            InitializeComponent();
        }

        private void bttn_View_Click(object sender, EventArgs e)
        {
            FillGrid();
        }


        private void FillGrid()
        {
            dgv_PTOMesurement.Rows.Clear();
            int i = 0;
            foreach (DataRow dt in _bal.getPtoMeasurementData(dtp_StartDate.Text, dtp_endDate.Text).Rows)
            {
                int startcompare = DateTime.Compare(Convert.ToDateTime(Convert.ToDateTime(dt["createdate"]).ToString("dd/MM/yyyy")), Convert.ToDateTime(dtp_StartDate.Text));
                int endcompare = DateTime.Compare(Convert.ToDateTime(Convert.ToDateTime(dt["createdate"]).ToString("dd/MM/yyyy")), Convert.ToDateTime(dtp_endDate.Text));
                if ((startcompare >= 0) && endcompare <= 0)
                {
                    dgv_PTOMesurement.Rows.Add();
                    dgv_PTOMesurement.Rows[i].Cells["modelid"].Value = _bal.getModel(Convert.ToInt32(dt["modelid"].ToString())).Rows[0]["model_name"].ToString();
                    dgv_PTOMesurement.Rows[i].Cells["master_value"].Value = dt["master_value"];
                    dgv_PTOMesurement.Rows[i].Cells["dialguage_reading"].Value = dt["dialguage_reading"];
                    dgv_PTOMesurement.Rows[i].Cells["actual_reading"].Value = dt["actual_reading"];
                    dgv_PTOMesurement.Rows[i].Cells["manual_reading"].Value = dt["manual_reading"];
                    dgv_PTOMesurement.Rows[i].Cells["shim_required"].Value = dt["shim_required"];
                    dgv_PTOMesurement.Rows[i].Cells["total_shims"].Value = dt["total_shims"];
                    dgv_PTOMesurement.Rows[i].Cells["shim_required_verification"].Value = dt["shim_required_verification"];
                    dgv_PTOMesurement.Rows[i].Cells["operator_name"].Value = dt["operator_name"];
                    dgv_PTOMesurement.Rows[i].Cells["createdate"].Value = dt["createdate"];
                    i++;
                }
            }
        }

        private void Frm_DatewiseReport_Load(object sender, EventArgs e)
        {
            dgv_PTOMesurement.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 12F, FontStyle.Bold);
        }

        private void bttn_exportexcel_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            //Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet1;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //xlWorkSheet1 = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            int i = 0;
            int j = 0;
            for (i = 0; i < 1; i++)
            {
                for (j = 0; j <= dgv_PTOMesurement.ColumnCount - 1; j++)
                {
                    string header = dgv_PTOMesurement.Columns[j].HeaderText;
                    xlWorkSheet.Cells[i + 1, j + 1] = header;
                }
            }

            for (i = 0; i <= dgv_PTOMesurement.RowCount - 1; i++)
            {
                for (j = 0; j <= dgv_PTOMesurement.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dgv_PTOMesurement[j, i];
                    xlWorkSheet.Cells[i + 2, j + 1] = cell.Value;
                }
            }
            // xlWorkSheet1.Cells[dgv_problem_master.RowCount, dgv_problem_master.ColumnCount] = "http://csharp.net-informations.com";
            //xlWorkSheet1.Cells[dgv_problem_master.RowCount, dgv_problem_master.ColumnCount] = "Chart Image";

            //xlWorkSheet1.Shapes.AddPicture(@"F:\vaishali\LearningMOBI\WebProject\img1.jpg", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 50, 50 + dgv_problem_master.RowCount, 300, 45); 
            string time = DateTime.Now.ToString("HH:mm").Replace(':', '_');
            string excelsheet_name = Environment.CurrentDirectory + "\\ExcelFiles\\" + "Spline_Shaft_Measurment" + "_From_" + Convert.ToDateTime(dtp_StartDate.Text).ToString("dd-MM-yyyy") + "_To_" + Convert.ToDateTime(dtp_StartDate.Text).ToString("dd-MM-yyyy")+"_"+time + ".xls";
            xlWorkBook.SaveAs(excelsheet_name, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            // releaseObject(xlWorkSheet1);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file " + excelsheet_name);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

    }
}
