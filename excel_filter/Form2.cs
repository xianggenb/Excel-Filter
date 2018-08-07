using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace excel_filter
{
    public partial class Form2 : Form
    {
        public  string form_num;

        public Form2()
        {
            InitializeComponent();
        }
        
        private void Form2_Load(object sender, EventArgs e)
        {
            int x = 50;
            int y = 80;

            if (Form1.tableNames.Count > 1) {

                for (int i = 0; i < Form1.tableNames.Count; i++) {
                    RadioButton rdo = new RadioButton();
                    rdo.Name = Form1.tableNames[i];
                    rdo.Text = Form1.tableNames[i];
                    rdo.ForeColor = Color.Black;
                    rdo.Location = new Point(x, y);
                    x += 150;
                    if (x > 500) {
                        x = 50;
                        y += 60;
                    
                    
                    }

                    if (y > 300) {
                        MessageBox.Show("表太多了", "太多了");

                        return;
                    
                    
                    }

                    this.Controls.Add(rdo);
                
                
                }
            
            
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RadioButton rbSelected = this.Controls
                         .OfType<RadioButton>()
                         .FirstOrDefault(r => r.Checked);
            form_num = rbSelected.Name;
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;

        }
    }
}
