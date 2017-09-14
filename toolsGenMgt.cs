using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;



using Aspose.Cells; //已经添加了应用 
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace sunproject
{
    public partial class toolsGenMgt : Form
    {
        public toolsGenMgt()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog=new OpenFileDialog();
            
            openFileDialog.Filter = "Excel结果表格|*.xlsx|所有文件|*.*";
			openFileDialog.RestoreDirectory=true;
			openFileDialog.FilterIndex=1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = openFileDialog.FileName;
            }
        }
        public int roofNum;
        private void button3_Click(object sender, EventArgs e)
        {
            textBox4.Text = "*TH-GRAPH";
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(textBox2.Text);
            Aspose.Cells.Workbook fworkbook = new Aspose.Cells.Workbook(textBox8.Text);
            Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];//getSheetByName(workbook,"总信息");
            Aspose.Cells.Worksheet fsheet = fworkbook.Worksheets[0];//getSheetByName(workbook,"总信息");
            string currentLine = "a";
            int i = 1;
            List<string> ids = new List<string>();//BRB塑性铰IDS
            List<string> fids = new List<string>();//楼层塑性铰IDS
            while(currentLine!=""){
                if (textBox3.Text=="0" || textBox3.Text.IndexOf("["+sheet.Cells["F" + i.ToString()].StringValue+"]") > -1)
                    ids.Add(sheet.Cells["A" + i.ToString()].StringValue.ToString());
                i++;
                currentLine = sheet.Cells["A" + i.ToString()].StringValue.ToString();
            }
            currentLine = "a";
            i = 1;
            while (currentLine != "")//遍历楼层非塑性铰定义
            {
                var str = (fsheet.Cells["A" + i.ToString()].StringValue.ToString());

                if ((fsheet.Cells["B" + i.ToString()].StringValue.ToString())=="桁架单元")
                {
                    textBox7.Text += "\r\n" + "  0, 9, f" + str + "maxd, " + textBox5.Text + ", " + str + ", 1, 1, 0, 2";//0 2 桁架单元  1  0 梁单元
                    textBox7.Text += "\r\n" + "  0, 9, f" + str + "maxl, " + textBox5.Text + ", " + str + ", 2, 1, 0, 2";
                }
                else
                {
                    textBox7.Text += "\r\n" + "  0, 9, f" + str + "maxd, " + textBox5.Text + ", " + str + ", 1, 1, 1, 0";//0 2 桁架单元  1  0 梁单元
                    textBox7.Text += "\r\n" + "  0, 9, f" + str + "maxl, " + textBox5.Text + ", " + str + ", 2, 1, 1, 0";
                }

                i++;
                currentLine = fsheet.Cells["A" + i.ToString()].StringValue.ToString();
            }


            foreach (string str in ids)//遍历BRB IDS
            {
                if (radioButton1.Checked)
                {
                    textBox4.Text += "\r\n" + "  0, 9, " + str + "maxd, " + textBox5.Text + ", " + str + ", 1, 1, 0, 2";//0 2 桁架单元  1  0 梁单元
                    textBox4.Text += "\r\n" + "  0, 9, " + str + "maxl, " + textBox5.Text + ", " + str + ", 2, 1, 0, 2";
                }
                else
                {
                    textBox4.Text += "\r\n" + "  0, 9, " + str + "maxd, " + textBox5.Text + ", " + str + ", 1, 1, 1, 0";//0 2 桁架单元  1  0 梁单元
                    textBox4.Text += "\r\n" + "  0, 9, " + str + "maxl, " + textBox5.Text + ", " + str + ", 2, 1, 1, 0";
                }
            }
            label6.Text = "共："+ids.Count.ToString();

            for (i = 1; i <= roofNum; i++)//遍历 楼层 
            {
                textBox6.Text += "\r\n" + "  0, 12, " + i + "F, " + textBox5.Text + ", " + i + "F, 1, 1, 1";//0 2 桁架单元  1  0 梁单元
            }
        }
    }
}
