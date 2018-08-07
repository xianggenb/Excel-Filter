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

namespace excel_filter
{
    public partial class Form1 : Form
    {
        public static string type;
        public static string cal_type;
        public static double cal;
        private struct cell_infor {
            public dynamic filter_targer {set;get;}
            public int row_num { set; get; }
            public int col_num { set; get; }
        }
       public static  List<string> tableNames = new List<string>();
      // public static List<string> col_headers = new List<string>();
        cell_infor overall_cell = new cell_infor() ;
        
        public Form1()
        {
            
            InitializeComponent();
            button1.Enabled = true;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            synbtn.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.openFileDialog1.FileName = "";
            this.openFileDialog1.Title = "选择Excel文件";
            this.openFileDialog1.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            this.openFileDialog1.Filter = "excel 表格 (*.xlsx)|*.xlsx";
            this.openFileDialog1.FilterIndex = 1;
        
            string file = null;
            if (this.openFileDialog1.ShowDialog(this) == DialogResult.OK) { 


               file = openFileDialog1.FileName;
               System.Diagnostics.Debug.WriteLine(file);
               dataGridView1.DataSource = GetExcelTableByOleDB(file);
             
               for (int i = 0; i < dataGridView1.Rows.Count; i++) {

                   if (dataGridView1.Rows[i] == null) {
                       dataGridView1.Rows.RemoveAt(i);
                   
                   }
               
               }
                   dataGridView1.Update();

               button2.Enabled = true;
               button3.Enabled = true;
               button4.Enabled = true;
               button5.Enabled = true;
               button6.Enabled = true;
               button7.Enabled = true;
               synbtn.Enabled = true;

        
            }

            foreach (DataGridViewColumn c in dataGridView1.Columns)
            {

                checkedListBox1.Items.Add(c.HeaderText, CheckState.Checked);


            }
         //   this.checkedListBox1.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkedListBox1_ItemCheck);
           
            this.button1.Enabled = false;
            return;

        }
      
        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex < 0 || e.ColumnIndex < 0) {
                if (e.RowIndex < 0) {
                    System.Diagnostics.Debug.WriteLine(dataGridView1.Columns[e.ColumnIndex].HeaderText);
                   // dataGridView1.Columns[e.RowIndex].HeaderText;
                
                }
               
                return;
            
            
            }
            this.ContextMenuStrip = MainStrip;
            overall_cell = new cell_infor();
            overall_cell.col_num = -1;
            overall_cell.row_num = -1;
            overall_cell.filter_targer = null;
            overall_cell.filter_targer =(dynamic) this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            overall_cell.row_num = e.RowIndex;
            overall_cell.col_num = e.ColumnIndex;
            
        }
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e) {

             type = "";
            if (dataGridView1.CurrentCell.Value.GetType() == type.GetType())
            {
                if (dataGridView1.CurrentCell.Value != null)
                {
                    int cur_col = dataGridView1.CurrentCell.ColumnIndex;
                    type = dataGridView1.Columns[cur_col].HeaderText;
                    DialogResult dialogResult = MessageBox.Show("基准类为 " + type, "设置成功", MessageBoxButtons.OKCancel);
                    if (dialogResult == DialogResult.OK) {
                    this.dataGridView1.CellClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellClick);

                        MessageBox.Show("选择求和列", "选择类别");
                        this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView3_CellClick);
                     
                        return;
                    
                    }
                    else {

                        this.dataGridView1.CellClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellClick);
                        return; }
                }

              

            }
            this.dataGridView1.CellClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellClick);
            return;
        }
        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            cal_type = "";
            cal = 0;
            if (dataGridView1.CurrentCell.Value.GetType() == cal.GetType())
            {
                if (dataGridView1.CurrentCell.Value != null)
                {
                    int cur_col = dataGridView1.CurrentCell.ColumnIndex;
                    cal_type = dataGridView1.Columns[cur_col].HeaderText;
                    MessageBox.Show("求和列为 " + cal_type, "设置成功");

                    if (!string.IsNullOrWhiteSpace(type) && !string.IsNullOrWhiteSpace(cal_type))
                    {
                        int type_index = -2;
                        int sum_index = -2;
                        List<string> type_list = new List<string>();
                        Dictionary<String, Tuple<double, double>> result = new Dictionary<string, Tuple<double, double>>();
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            if (dataGridView1.Columns[i].HeaderText == type)
                            {
                                type_index = i;

                            }
                            if (dataGridView1.Columns[i].HeaderText == cal_type) {
                                sum_index = i;
                            
                            }

                        }
                        if (type_index > 0 || sum_index>0)
                        {

                            for (int j = 0; j < dataGridView1.Rows.Count; j++)
                            {
                                if (dataGridView1[type_index, j].Value != null)
                                {
                                    if (!type_list.Contains(dataGridView1[type_index, j].Value.ToString()))
                                    {

                                        type_list.Add(dataGridView1[type_index, j].Value.ToString());

                                    }
                                }
                            }
                        


                        }
                        foreach (string c in type_list) {
                            double sum = 0;
                            int freq = 0;
                            for (int k = 0; k < dataGridView1.Rows.Count-1; k++) {
                                if (dataGridView1[type_index, k].Value.ToString() == c) { 
                                
                                sum+=(double)dataGridView1[sum_index,k].Value;
                                freq++;
                                }

                               
                            } 
                            result.Add(c,new Tuple<double,double>(sum,freq));
                            System.Diagnostics.Debug.WriteLine(c+" "+sum+" "+freq );


                          
                        
                        }
                        using (child_form frmcf = new child_form())
                        {
                         
                            frmcf.res = result;
                            frmcf.s1 = dataGridView1.Columns[type_index].HeaderText;
                            frmcf.s2 = dataGridView1.Columns[sum_index].HeaderText;
                         //   dataGridView1 = cview;
                            frmcf.ShowDialog();

                        }
                        

                    }
                    else { return; }
                    this.dataGridView1.CellClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView3_CellClick);

                    return;
                }



            }
            this.dataGridView1.CellClick -= new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView3_CellClick);
            return;
        }
        public static DataTable GetExcelTableByOleDB(string strExcelPath)
        {
            string tableName=null;
            try
            {
                DataTable dtExcel = new DataTable();
                DataSet ds = new DataSet();
                string strExtension = System.IO.Path.GetExtension(strExcelPath);
                string strFileName = System.IO.Path.GetFileName(strExcelPath);
                OleDbConnection objConn = null;
                switch (strExtension)
                {
                    case ".xls":
                        objConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strExcelPath + ";" + "Extended Properties=\"Excel 8.0;HDR=yes;IMEX=1;\"");
                        break;
                    case ".xlsx":
                        objConn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strExcelPath + ";" + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"");//此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串)  备注： "HDR=yes;"是说Excel文件的第一行是列名而不是数，"HDR=No;"正好与前面的相反。"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。 
                        break;
                    default:
                        objConn = null;
                        break;
                }
                if (objConn == null)
                {
                    return null;
                }
                objConn.Open();
                DataTable userTables = null;
                string[] restrictions = new string[4];
                restrictions[3] = "Table";
                userTables = objConn.GetSchema("Tables", restrictions);
            
                for (int i = 0; i < userTables.Rows.Count; i++)
                {
                    if (userTables.Rows[i][2].ToString().EndsWith("$")) { 
                    tableNames.Add(userTables.Rows[i][2].ToString());
                    }
                }

                if(tableNames.Count>=2){
                    using (Form2 frmselct = new Form2()) {

                        if (frmselct.ShowDialog() == DialogResult.OK) {
                            tableName = frmselct.form_num;

                        
                        }
                    }
                
                }
                else if (tableNames.Count == 1) {
                    tableName = tableNames[0];
                
                
                }
                if (tableName != null)
                {
                    string strSql = "select * from ["+tableName+"]";
                    OleDbCommand objCmd = new OleDbCommand(strSql, objConn);
                    OleDbDataAdapter myData = new OleDbDataAdapter(strSql, objConn);
                    myData.Fill(ds, tableName);
                    objConn.Close();
                    dtExcel = ds.Tables[tableName];

                   
                }
          return dtExcel;


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "\r\n" + ex.StackTrace);
                return null;
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {  
            dataGridView1.AllowUserToAddRows = false;
           
           if (overall_cell.filter_targer != null || overall_cell.row_num > 0 || overall_cell.col_num > 0)
            {

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if ((dynamic)dataGridView1.Rows[i].Cells[overall_cell.col_num].Value != overall_cell.filter_targer)
                    {

                        dataGridView1.Rows.Remove(dataGridView1.Rows[i]);
                        i--;
                    }

                }



            }
            
            else
            {

                MessageBox.Show("未选中目标", "错误");
                return;

            }
           dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            if (overall_cell.filter_targer != null || overall_cell.row_num > 0 || overall_cell.col_num > 0)
            {

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if ((dynamic)dataGridView1.Rows[i].Cells[overall_cell.col_num].Value == overall_cell.filter_targer)
                    {

                        dataGridView1.Rows.Remove(dataGridView1.Rows[i]);
                        i--;
                    }

                }



            }
            else
            {

                MessageBox.Show("未选中目标", "错误");
                return;

            }
            dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            save();
            return;
        }

        private void synbtn_Click(object sender, EventArgs e)
        {
            for (int k = 0; k < checkedListBox1.Items.Count; k++) {

                if (!checkedListBox1.CheckedItems.Contains(checkedListBox1.Items[k])) {
                
                     for (int i = 0; i < dataGridView1.Columns.Count; i++) {

                         if (dataGridView1.Columns[i].HeaderText == (string)checkedListBox1.Items[k]) {
                             dataGridView1.Columns.Remove(dataGridView1.Columns[i]);
                         
                         }
            
            
            
            }
                        checkedListBox1.Items.Remove(checkedListBox1.Items[k]); 
                       k--;
               
                
                
                }
            
            }
            checkedListBox1.Update();

         
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell == null)
            {
                return;

            }
            else {
                int count = 0;
                int cur_col = -2;
                int sum = 0;
                MessageBox.Show("设置基准求和类","选择类别");
                this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView2_CellClick);
               
                cur_col = dataGridView1.CurrentCell.ColumnIndex;
                DataGridViewColumn operated_col = dataGridView1.Columns[cur_col];
              

            }

        }
        private void selected_dbclick(object sender, MouseEventArgs e)
        { 
        string type=null;
        if (dataGridView1.CurrentCell.Value.GetType() == type.GetType()) {
            if (dataGridView1.CurrentCell.Value != null) {
                type = dataGridView1.CurrentCell.Value.ToString();
                MessageBox.Show("基准类为+ "+type,"设置成功");
            }

        
        
        }
        this.MouseDoubleClick -= new MouseEventHandler(this.selected_dbclick);
        
        }
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

          

        }

        private void button6_Click(object sender, EventArgs e)
        {


            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, true);
            }
            return;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
            return;
        }
        public void save()
        {
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

            for (int i = 0; i < dataGridView1.ColumnCount; i++)

            { worksheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText; }

            for (int r = 0; r < dataGridView1.Rows.Count; r++)
            {
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {

                    worksheet.Cells[r + 2, i + 1] = dataGridView1.Rows[r].Cells[i].Value;

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
