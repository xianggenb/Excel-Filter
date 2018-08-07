using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using Microsoft.Office;

namespace excel_filter
{
    public partial class child_form : Form
    {
        public child_form()
        {
         
            InitializeComponent();
      
        }
        public DataGridView dgv {get;set; }
        public Dictionary<String, Tuple<double, double>> res{get;set;}
        public String s1 { get; set; }
        public String s2 { get; set; }
        private void button1_Click(object sender, EventArgs e)
        {
          /*  String strExcelPath = null;

            this.saveFileDialog1.FileName = "";
            this.saveFileDialog1.Title = "保存excel文件";
            this.saveFileDialog1.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            this.saveFileDialog1.Filter = "excel 表格 (*.xlsx)|*.xlsx";
            this.saveFileDialog1.FilterIndex = 1;

            if (this.saveFileDialog1.ShowDialog(this) == DialogResult.OK)
            {


                strExcelPath = saveFileDialog1.FileName;
              
            }
            ToExcel(this.dataGridView2,strExcelPath);

            */
            save();
            return;
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void child_form_Load(object sender, EventArgs e)
        {
           
            dataGridView2.ColumnCount = 3;
            dataGridView2.Columns[0].Name = s1;
            dataGridView2.Columns[1].Name = s2;
            dataGridView2.Columns[2].Name = "频次";
            foreach (KeyValuePair<string, Tuple<double, double>> c in res)
            {
                dataGridView2.Rows.Add(c.Key, c.Value.Item1, c.Value.Item2);


            }
            dataGridView2.Update();
        
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
      
        public void save() {
            string fileName = "";

            string saveFileName = "";



            saveFileDialog1.DefaultExt = "xlsx";

            saveFileDialog1.Filter = "Excel文件|*.xlsx";

            saveFileDialog1.FileName = fileName;

            saveFileDialog1.ShowDialog();

            saveFileName = saveFileDialog1.FileName;

            if (saveFileName.IndexOf(":") < 0) return;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {

                MessageBox.Show("无法创建Excel对象，您的电脑可能未安装Excel");

                return;

            }

            Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;

            Microsoft.Office.Interop.Excel.Workbook workbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);

            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Worksheets[1];//取得sheet1 

            for (int i = 0; i < dataGridView2.ColumnCount; i++)

            { worksheet.Cells[1, i + 1] = dataGridView2.Columns[i].HeaderText; }

            for (int r = 0; r < dataGridView2.Rows.Count; r++)
            {
                for (int i = 0; i < dataGridView2.ColumnCount; i++)
                {

                    worksheet.Cells[r + 2, i + 1] = dataGridView2.Rows[r].Cells[i].Value;

                }

                System.Windows.Forms.Application.DoEvents();

            }

            worksheet.Columns.EntireColumn.AutoFit();

            MessageBox.Show(fileName + "资料保存成功", "提示", MessageBoxButtons.OK);

            if (saveFileName != "")
            {

                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(saveFileName);     
                }

                catch (Exception ex)
                {                    

                    MessageBox.Show("导出文件时出错,文件可能正被打开！\n" + ex.Message);

                }

            }
            xlApp.Quit();
            GC.Collect();
        
        }
    }
}
