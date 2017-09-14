/*
 * Created by SharpDevelop.
 * User: Yangxin
 * Date: 2015/1/15
 * Time: 21:30
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

using System.Windows.Forms.DataVisualization.Charting;
//using System.Web.UI.DataVisualization.Charting;

using System.Data;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace sunproject
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
    /// 
	public partial class MainForm : Form
	{
		
        public Thread thread = null;
        //public Dictionary<string, Dictionary<int, Dictionary<string, double>>> floorStaus = new Dictionary<string, Dictionary<int, Dictionary<string, double>>>();

        [STAThread]
		public static void Main(string[] args)
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);
			Application.Run(new MainForm());
		}

        //***********************
        //protected delegate void UpdateControlText(string strText); //定义一个委托 
        //定义更新控件的方法  
        //protected void updateControlText(string line){this.listBox2.Items.Add(line.ToString());}

        protected delegate void ifendsat(bool staus); 
        protected void ifend(bool staus)
        {
            listBox2.SelectedIndex = listBox2.Items.Add("读取Midas结果完成[设定基本参数后进行下一步计算]");

            //button2.Enabled = true;
            //button3.Enabled = true;
            button4.Enabled = true;
            toolStripProgressBar1.Visible = false;
            toolStripStatusLabel1.Text = "计算完成";
            dataGridView2.Rows[1].Cells[1].Value = Global.sec.ToString("F2");
            dataGridView2.Rows[2].Cells[1].Value = Global.MidasModelCalTime.ToString("F2");
            textBox3.Text = Global.MidasModelCalTime.ToString("F2");
            
        }
        /*
        protected void reinit(bool staus)
        {
            button2.Enabled = true;
            toolStripProgressBar1.Visible = false;
            toolStripStatusLabel1.Text = "计算完成";
        }*/

		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();

            int index;
            dataGridView2.Rows[dataGridView2.Rows.Add()].Cells[0].Value = "fv";
            dataGridView2.Rows[0].Cells[1].Value = "235";
            dataGridView2.Rows[0].Cells[2].Value = "BRB屈服强度";

            dataGridView2.Rows[dataGridView2.Rows.Add()].Cells[0].Value = "时间步长(s)";
            dataGridView2.Rows[1].Cells[1].Value = Global.sec.ToString("F2");
            dataGridView2.Rows[1].Cells[2].Value = "自动获取";

            dataGridView2.Rows[dataGridView2.Rows.Add()].Cells[0].Value = "计算时长(s)";
            dataGridView2.Rows[2].Cells[1].Value = "20";
            dataGridView2.Rows[2].Cells[2].Value = "自动获取";//暂不支持局部时段计算

            dataGridView2.Rows[dataGridView2.Rows.Add()].Cells[0].Value = "BRB等效截面积";
            dataGridView2.Rows[3].Cells[1].Value = "1";
            dataGridView2.Rows[3].Cells[2].Value = "按'截面信息.xls'自动获取";

            dataGridView2.Rows[dataGridView2.Rows.Add()].Cells[0].Value = "生成计算书";
            dataGridView2.Rows[4].Cells[1].Value = "1";
            dataGridView2.Rows[4].Cells[2].Value = "计算完生成EXCEL计算书";
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}

        public bool debugMode = false;
        public string fileName, filePath;
        public Dictionary<string, string> filesFv = new Dictionary<string, string>();
        public Dictionary<string, string> filesU = new Dictionary<string, string>();
        public Dictionary<string, string> filesJ = new Dictionary<string, string>();
        public List<string> filesX = new List<string>();
            
        public db db = new db();

		void Button1Click(object sender, EventArgs e)
		{
			OpenFileDialog openFileDialog=new OpenFileDialog();
            
			//openFileDialog.InitialDirectory="c:\\";
            openFileDialog.Filter = "Midas工程文件|*.mgb|所有文件|*.*";
			openFileDialog.RestoreDirectory=true;
			openFileDialog.FilterIndex=1;
			if (openFileDialog.ShowDialog()==DialogResult.OK)
			{
				this.textBox1.Text=openFileDialog.FileName;
                fileName = openFileDialog.SafeFileName;
                filePath = openFileDialog.FileName.Replace(openFileDialog.SafeFileName, "");

                filePathMethor(filePath);
                return;

                DirectoryInfo di = new DirectoryInfo(filePath);
                FileInfo[] files = di.GetFiles();
                //this.label11.Text = "";
                foreach (FileInfo eachF in files)
                {
                    if (eachF.Name.IndexOf("层位移") > -1)
                    {
                        string name = eachF.Name.Replace(".txt","").Split(new char[1] { ' ' })[0];
                        if (!this.listBox3.Items.Contains(name))
                            this.listBox3.Items.Add(name);
                        filesU.Add(name, eachF.Name);
                    }
                    if (eachF.Name.IndexOf("层剪力") > -1)
                    {
                        string name = eachF.Name.Replace(".txt", "").Split(new char[1] { ' ' })[0];
                        if (!this.listBox3.Items.Contains(name))
                            this.listBox3.Items.Add(name);
                        filesFv.Add(name,eachF.Name);
                    }
                    if (eachF.Name.IndexOf("塑性铰") > -1)
                    {
                        string name = eachF.Name.Replace(".txt", "").Split(new char[1] { ' ' })[0];
                        if (!this.listBox3.Items.Contains(name))
                            this.listBox3.Items.Add(name);
                        filesJ.Add(name, eachF.Name);
                    }
                }
                if (File.Exists(filePath + "//excel"))
                {
                    DirectoryInfo diXls = new DirectoryInfo(filePath + "//excel");
                    FileInfo[] filesXls = diXls.GetFiles();
                    foreach (FileInfo eachF in filesXls)
                    {
                        string name = eachF.Name.Replace(".xls", "");
                        filesX.Add(eachF.Name);
                    }
                }
			}
		}
        private void filePathMethor(string filePath)
        {
            this.textBox1.Text = filePath;
            DirectoryInfo di = new DirectoryInfo(filePath);
            FileInfo[] files = di.GetFiles();
            //this.label11.Text = "";
            foreach (FileInfo eachF in files)
            {
                if (eachF.Name.IndexOf("层位移") > -1)
                {
                    string name = eachF.Name.Replace(".txt", "").Split(new char[1] { ' ' })[0];
                    if (!this.listBox3.Items.Contains(name))
                        this.listBox3.Items.Add(name);
                    filesU.Add(name, eachF.Name);
                }
                if (eachF.Name.IndexOf("层剪力") > -1)
                {
                    string name = eachF.Name.Replace(".txt", "").Split(new char[1] { ' ' })[0];
                    if (!this.listBox3.Items.Contains(name))
                        this.listBox3.Items.Add(name);
                    filesFv.Add(name, eachF.Name);
                }
                if (eachF.Name.IndexOf("塑性铰") > -1)
                {
                    string name = eachF.Name.Replace(".txt", "").Split(new char[1] { ' ' })[0];
                    if (!this.listBox3.Items.Contains(name))
                        this.listBox3.Items.Add(name);
                    filesJ.Add(name, eachF.Name);
                }
            }
            if (File.Exists(filePath + "//excel"))
            {
                DirectoryInfo diXls = new DirectoryInfo(filePath + "//excel");
                FileInfo[] filesXls = diXls.GetFiles();
                foreach (FileInfo eachF in filesXls)
                {
                    string name = eachF.Name.Replace(".xls", "");
                    filesX.Add(eachF.Name);
                }
            }
            return;
        }
		void Button2Click(object sender, EventArgs e)
		{
            if (this.listBox3.Text == "")
            {
                MessageBox.Show("未选择地震波");
                return;
            }

            debugMode = checkBox1.Checked;
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            dataGridView1.Rows.Clear();
            if(!chart1.Series.IsUniqueName("0"))
                chart1.Series[0].Points.Clear();

            toolStripProgressBar1.Visible = true;
            //toolStripProgressBar1
            button1.Enabled = false;
            button2.Enabled = false;
            //button3.Enabled = false;
            button4.Enabled = false;
            textBox5.Enabled = false;
            textBox6.Enabled = false;
            checkBox1.Enabled = false;
            checkBox3.Enabled = false;

            toolStripStatusLabel1.Text="计算中...";
            this.listBox3.Enabled = false;
            
            thread = new Thread(new ThreadStart(threadGetFile));
            thread.Start();
		}

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (thread != null) thread.Abort();
        }
        public delegate void TaskDelegate();
        private delegate void InvokeMethodDelegate(Control control, TaskDelegate handler);

        public static void SafeInvoke(Control control, TaskDelegate handler)
        {
            if (control.InvokeRequired)
            {
                while (!control.IsHandleCreated)
                {
                    if (control.Disposing || control.IsDisposed)
                        return;
                }
                IAsyncResult result = control.BeginInvoke(new InvokeMethodDelegate(SafeInvoke), new object[] { control, handler });
                control.EndInvoke(result);//获取委托执行结果的返回值
                return;
            }
            IAsyncResult result2 = control.BeginInvoke(handler);
            control.EndInvoke(result2);
        }


        public void showDate(bool clear = false,string set = "FX")
        {
            
            double umax = 0; double umin = 0;//double fmax,fmin;
            bool newDataSet = true;
            this.dataGridView1.Rows.Clear();
            if (clear) this.chart1.Series.Clear();
            else
            {
                if (radioButton2.Checked) set = "FY";
                else if (radioButton3.Checked) set = "UX";
                else if (radioButton4.Checked) set = "UY";
                else if (radioButton5.Checked) set = "Ws";
                else if (radioButton6.Checked) set = "滞回曲线";
                else if (radioButton7.Checked) set = "znb";
                else if (radioButton8.Checked) set = "znb25";
            }

            if (listBox1.SelectedIndex == -1) return;
            string id = listBox1.Items[listBox1.SelectedIndex].ToString().Split(new char[1] { '|' })[0];

            if (chart1.Series.IsUniqueName(id.ToString() + "-" + set))
                chart1.Series.Add(id.ToString() + "-" + set);
            Series ser = chart1.Series.FindByName(id.ToString() + "-" + set);
            ser.Points.Clear();
            ser.ChartType = SeriesChartType.Line;
            double total = 0, preTotal = 0;
            int totalAmount = 0 , preTotalAmount = 0;
            //dbMathPoint d1 = new dbMathPoint(0, 0);
            //dbMathPoint d2 = new dbMathPoint(0, 0);
            //dbMathPoint d3 = new dbMathPoint(0, 0);
            //dbMathPoint d4 = new dbMathPoint(0, 0);

            //包络计算方法数据显示
            //label13.Text = "当前层MaxFiT="+db.status["CHiCHi"].floors[id].maxFiT.ToString();
            //label16.Text = "当前层MaxUiT=" + db.status["CHiCHi"].floors[id].maxUiT.ToString();
            double maxBlTime = dbMathPoint.zhqxEdge(
                db.status["CHiCHi"].floors[id].timeStat,
                Convert.ToDouble( textBox3.Text ));//包络计算滞回曲线外包轮廓的max时刻


            foreach (KeyValuePair<double, dbPointStatus> pDb in db.status["CHiCHi"].floors[id].timeStat)
            {
                double selectedData = pDb.Value.fx;
                if (set == "FY") selectedData = pDb.Value.fy;
                else if (set == "UX") selectedData = pDb.Value.ux;
                else if (set == "UY") selectedData = pDb.Value.uy;
                else if (set == "Ws") selectedData = pDb.Value.ws;//pDb.Value.ux * pDb.Value.fx;
                //else if (set == "Ws" && (gks.Value.ifFloor)) selectedData = pDb.Value.ux * pDb.Value.fx;
                else if (set == "滞回曲线") selectedData = pDb.Value.fx;
                else if (set == "znb" || set == "znb25")
                {
                    selectedData =
                    db.status["CHiCHi"].floors["消能器Total"].timeStat[pDb.Key].ws / (4 * 3.1415926535 *
                    db.status["CHiCHi"].floors["楼层Total"].timeStat[pDb.Key].ws) * 100;
                    if (selectedData < 0) selectedData = -selectedData;
                    if (set == "znb25" ){
                        if(Math.Abs(selectedData) > 25) selectedData = 25;
                        selectedData = Math.Abs(selectedData);
                    }
                }
                if ((id == "消能器Total" || id == "楼层Total") && set == "Ws") selectedData = pDb.Value.ws;//汇总后的数据

                //if (set != "znb25"|| selectedData > 0)  取平均方法
                {
                    total += selectedData;
                    totalAmount++;
                }
                if (newDataSet) { newDataSet = false; umin = selectedData; umax = selectedData; }
                if (pDb.Key > Convert.ToDouble(textBox2.Text) && pDb.Key < Convert.ToDouble(textBox3.Text))
                {
                    //if (selectedData > 0)   //取平均方法
                    {
                        preTotal += selectedData;
                        preTotalAmount++;
                    }

                    if (selectedData < umin) umin = selectedData;
                    else if (selectedData > umax) umax = selectedData;
                    if (set == "滞回曲线")
                    {
                        ser.Points.AddXY(pDb.Value.ux, pDb.Value.fx);//Max-D  //插入滞回曲线点集合
                    }
                    else
                        ser.Points.AddXY(pDb.Key, selectedData);//Max-D

                    int drId = dataGridView1.Rows.Add();
                    dataGridView1.Rows[drId].Cells[0].Value = pDb.Key;
                    dataGridView1.Rows[drId].Cells[1].Value = selectedData;
                    //dataGridView1.Rows[drId].Cells[3].Value = "";
                }
            }
            if (
                (set == "滞回曲线" && 
                (id != "消能器Total" && id != "楼层Total") 
                && !db.status["CHiCHi"].floors[id].ifFloor)
                ||
                set == "znb25"
                )
            MainForm.SafeInvoke(this.label7, new TaskDelegate(delegate()
            {
                label7.Text = ("累加值:"+total.ToString("E3")+" 平均值:"+(total/totalAmount).ToString("F2"));
                //if (set == "滞回曲线")
                    //label7.Text = "124点平行四边形面积:"+dbMathPoint.sBy3P(d1, d2, d4).ToString("E2");
            }));
            MainForm.SafeInvoke(this.label20, new TaskDelegate(delegate()
            {
                label20.Text = ("时段平均:" + (preTotal / preTotalAmount).ToString("F2"));
                //if (set == "滞回曲线")
                //label7.Text = "124点平行四边形面积:"+dbMathPoint.sBy3P(d1, d2, d4).ToString("E2");
            }));

            if (id == "消能器Total" && set == "Ws") 
                label14.Text = "消能器耗能： 总[ "+total.ToString("E2")+" ] MAX. [ "+ umax.ToString("E2")+" ]";
            if (id == "楼层Total" && set == "Ws")
                label9.Text = "楼层耗能： 总[ " + total.ToString("E2") + " ] MAX. [ " + umax.ToString("E2") + " ]";
            if (set == "znb25")
            {
                label15.Text = "附加阻尼比： " + (total / totalAmount).ToString("F2") + " %";
                Global.resultZnbSb = (total / totalAmount);
            }

            MainForm.SafeInvoke(this.label8, new TaskDelegate(delegate()
            {
                label8.Text = (umin.ToString("E3") + "/" + umax.ToString("E3"));
            }));

            if (set == "滞回曲线")
            {
                listBox2.Items.Add(db.status["CHiCHi"].floors[id].tanxinD1.ux + "/" + db.status["CHiCHi"].floors[id].tanxinD1.fx);
                listBox2.Items.Add(db.status["CHiCHi"].floors[id].tanxinD2.ux + "/" + db.status["CHiCHi"].floors[id].tanxinD2.fx);
                MainForm.SafeInvoke(this.chart1, new TaskDelegate(delegate()  //添加滞回曲线2个弹性点范围
                {
                    if (chart1.Series.IsUniqueName(id.ToString() + "-" +  "弹性阶段"))
                        chart1.Series.Add(id.ToString() + "-" +  "弹性阶段");
                    Series ser1 = chart1.Series.FindByName(id.ToString() + "-" +  "弹性阶段");
                    ser1.Points.Clear();
                    ser1.MarkerSize = 10;
                    ser1.ChartType = SeriesChartType.Line;
                    int p1 = ser1.Points.AddXY(db.status["CHiCHi"].floors[id].tanxinD1.ux, db.status["CHiCHi"].floors[id].tanxinD1.fx);
                    int p2 = ser1.Points.AddXY(db.status["CHiCHi"].floors[id].tanxinD2.ux, db.status["CHiCHi"].floors[id].tanxinD2.fx);

                }));
            }
            MainForm.SafeInvoke(this.label18, new TaskDelegate(delegate(){
                this.label18.Text = "弹性阶段最大变形 = "+db.status["CHiCHi"].floors[id].tanxingD.ToString("F1");
            }));

            if (set == "滞回曲线")
            MainForm.SafeInvoke(this.chart1, new TaskDelegate(delegate()  //添加滞回曲线4个包络点集
            {
                if (chart1.Series.IsUniqueName(id.ToString() + "-" + set+"包络"))
                    chart1.Series.Add(id.ToString() + "-" + set + "包络");
                Series ser1 = chart1.Series.FindByName(id.ToString() + "-" + set + "包络");
                ser1.Points.Clear();
                ser1.ChartType = SeriesChartType.Line;
                ser1.Points.AddXY(db.status["CHiCHi"].floors[id].d0.ux, db.status["CHiCHi"].floors[id].d0.fx);
                ser1.Points.AddXY(db.status["CHiCHi"].floors[id].d1.ux, db.status["CHiCHi"].floors[id].d1.fx);
                ser1.Points.AddXY(db.status["CHiCHi"].floors[id].d2.ux, db.status["CHiCHi"].floors[id].d2.fx);
                ser1.Points.AddXY(db.status["CHiCHi"].floors[id].d3.ux, db.status["CHiCHi"].floors[id].d3.fx);
                ser1.Points.AddXY(db.status["CHiCHi"].floors[id].d0.ux, db.status["CHiCHi"].floors[id].d0.fx);
                ser1.BorderWidth = 2;
            }));
            MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
            {
                //listBox2.Items.Add (d1.x.ToString("E3") + "/" + d1.y.ToString("E3"));
                //listBox2.Items.Add(d2.x.ToString("E3") + "/" + d2.y.ToString("E3"));
                //listBox2.Items.Add(d3.x.ToString("E3") + "/" + d3.y.ToString("E3"));
                listBox2.Items.Add(db.status["CHiCHi"].floors[id].k.ToString() + "/" + db.status["CHiCHi"].floors[id].maxy.ToString());
            }));

        }
        public Dictionary<int, double> meterial2Area = new Dictionary<int, double>();
        public Dictionary<string, int> elementId2meterial = new Dictionary<string, int>();
        private void threadGetFile()
        {
            StreamReader sr = null;
            String line = "";
            int i2 = 0; int i3 = 0;
            //int countGK = 0;
            int countFloor = 1;
            string currentGK = "";
            Dictionary<int, Dictionary<string, double>> ps = new Dictionary<int, Dictionary<string, double>>();
            dbStatus dbs = new dbStatus();
            int i1 = 0;
            string currentSetName = "";
            string currentSetDir = "";
            double sec = -1;
            string thisComboBox1Text = "";
            MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
            {
                listBox2.Items.Add("读取工程文件基本参数");
                thisComboBox1Text = this.listBox3.Text;
            }));

            #region PKPM数据
            /*
            MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
            {
                listBox2.Items.Add("读取WDISP.OUT结果");
            }));

            floorStaus = new Dictionary<string, Dictionary<int, Dictionary<string, double>>>();
            if (File.Exists(filePath + "WDISP.OUT"))
                sr = new StreamReader(filePath + "WDISP.OUT", Encoding.Default);
            if(sr!=null)
            while ((line = sr.ReadLine()) != null)
            {
                if (i2 ==0 && line.IndexOf("楼层最大位移") > -1 && line.IndexOf("工况") > -1)
                {
                    //加新工况
                    i2 = 1;
                    countFloor = 0;
                    ps = new Dictionary<int, Dictionary<string, double>>();
                    dbs = new dbStatus();
                    countGK++;
                    currentGK = line;
                }
                if (i2 > 0 && ((line.Length==0&&countFloor>0) || line.IndexOf("方向最大层间位移角") > -1 || line.IndexOf("最大位移与层平均位移的比值") > -1))
                {
                    //完成一工况
                    i2 = 0;
                    i3 = 0;
                    countFloor = 1;
                    floorStaus.Add(countGK.ToString(), ps);
                    //db.status.Add(dbs);
                    MainForm.SafeInvoke(this.listBox1, new TaskDelegate(delegate()
                    {
                        listBox1.Items.Add(countGK.ToString() + "|  " + currentGK);
                    }));
                    ps = new Dictionary<int, Dictionary<string, double>>();
                }
                if (i2 > 0)
                {
                    line = Regex.Replace(line, @"\s+", " ");
                    string[] dZx = line.Split(new char[1] { ' ' });
                    if (dZx.Length < 5) continue;
                    Dictionary<string, double> pss = new Dictionary<string, double>();
                    if(i3>2 && i3%2==1){
                        //+一节点
                        pss.Add("Max-D", Convert.ToDouble(dZx[4]));
                        //+一层
                        ps.Add(countFloor, pss);

                        countFloor++;
                    }
                    i3++;
                }
            }

            MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
            {
                listBox2.Items.Add("读取WZQ.OUT结果");
            }));

            if (File.Exists(filePath + "WZQ.OUT"))
                sr = new StreamReader(filePath + "WZQ.OUT", Encoding.Default);
            //else return;
             //line;

            if(sr!=null)
            while ((line = sr.ReadLine()) != null)
            {
                MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                {
                    listBox2.Items.Add(line.ToString());
                }));
                
                if (i1 > 0)
                {
                    line = Regex.Replace(line, @"\s+", " ");
                    string[] dZx = line.Split(new char[1] { ' ' });
                    if (dZx.Length < 2) { i1 = 0; continue; ;}

                    MainForm.SafeInvoke(this.dataGridView1, new TaskDelegate(delegate()
                    {
                        int drId = dataGridView1.Rows.Add();
                        dataGridView1.Rows[drId].Cells[0].Value = i1.ToString();
                        dataGridView1.Rows[drId].Cells[1].Value = dZx[2].ToString();
                    }));

                    MainForm.SafeInvoke(this.chart1, new TaskDelegate(delegate()
                    {
                        if(!chart1.Series.IsUniqueName("0"))
                        chart1.Series[0].Points.AddXY(i1, dZx[2]);
                    }));

                    //int drId = dataGridView1.Rows.Add();
                    //dataGridView1.Rows[drId].Cells[0].Value = i1;
                    //dataGridView1.Rows[drId].Cells[1].Value = dZx[2].ToString();
                    i1++;
                }
                else if (i1 == 0 && line.IndexOf("振型号") > -1 && line.IndexOf("扭转系数") > -1)
                {
                    i1 = 1;
                    MainForm.SafeInvoke(this.dataGridView1, new TaskDelegate(delegate()
                    {
                        int drId = dataGridView1.Rows.Add();
                        dataGridView1.Rows[drId].Cells[0].Value = "振型";
                        dataGridView1.Rows[drId].Cells[1].Value = "周期";
                    }));
                }
            }
            */
            #endregion

            //读取截面信息
            Aspose.Cells.Workbook workbook = null;
            Aspose.Cells.Worksheet sheet = null;

            if (File.Exists(filePath + "截面信息.xlsx"))
            {
                workbook = new Aspose.Cells.Workbook(filePath + "截面信息.xlsx");
                sheet = workbook.Worksheets[0];
                string currentLine = "aaa";
                int iXlsLine = 1;
                while (currentLine != "")
                {
                        meterial2Area.Add(
                            Convert.ToInt32(sheet.Cells["A" + iXlsLine.ToString()].StringValue.ToString()),
                            Convert.ToDouble(sheet.Cells["E" + iXlsLine.ToString()].StringValue.ToString())
                            );
                    iXlsLine++;
                    currentLine = sheet.Cells["A" + iXlsLine.ToString()].StringValue.ToString();
                }
                MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                {
                    listBox2.Items.Add("读取截面信息[截面信息.xlsx]"+meterial2Area.Count.ToString());
                }));
            }
            if (File.Exists(filePath + "单元信息.xlsx"))
            {
                workbook = new Aspose.Cells.Workbook(filePath + "单元信息.xlsx");
                sheet = workbook.Worksheets[0];
                string currentLine = "aaa";
                int iXlsLine = 1;
                while (currentLine != "")
                {
                        elementId2meterial.Add(
                            sheet.Cells["A" + iXlsLine.ToString()].StringValue.ToString(),
                            Convert.ToInt32(sheet.Cells["F" + iXlsLine.ToString()].StringValue.ToString())
                            );
                    iXlsLine++;
                    currentLine = sheet.Cells["A" + iXlsLine.ToString()].StringValue.ToString();
                }
                MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                {
                    listBox2.Items.Add("读取截面信息[单元信息.xlsx]" + elementId2meterial.Count.ToString());
                }));
            }



            ///////////////////////////

            dbFloorStatus dbf = new dbFloorStatus();
            //**************************
            MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
            {
                listBox2.Items.Add("读取时程[CHiCHi]分析结果[Fv]");
            }));

            if (File.Exists(filePath + filesFv[thisComboBox1Text]))
                sr = new StreamReader(filePath + filesFv[thisComboBox1Text], Encoding.Default);
            if (sr != null)
            {
                i1 = 0;
                db.status.Clear();
                db.status.Add("CHiCHi",new dbStatus());
                
                while ((line = sr.ReadLine()) != null)
                {
                    if(!debugMode)
                    MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                    {
                        int selIndex = listBox2.Items.Add(line.ToString());
                        if (i1 == 0)
                          listBox2.SelectedIndex = (selIndex);
                    }));
                    
                    if (i1 == 0 && line.IndexOf("**") > -1 && line.IndexOf("midas ") < 0)
                    {
                        countFloor = 1;
                        i1 = 1;
                        //ps = new Dictionary<int, Dictionary<string, double>>();
                        //dbf = new dbFloorStatus();
                        string[] dddddd = Regex.Replace(line, @"\s+", " ").Split(new char[1] { ' ' })[1].Split(new char[1] { '-' });
                        currentSetName = dddddd[0].Replace("F","");//+dddddd[1]
                        currentSetDir = dddddd[1].Replace("-","");
                        if (currentSetDir == "Y" && checkBox3.Checked) currentSetDir = "X";
                        else if (currentSetDir == "X" && checkBox3.Checked) currentSetDir = "Y";
                        if (currentSetName == "wd") currentSetName = "Roof";
                        if (currentSetName == "屋顶") currentSetName = "Roof";
                        if (textBox6.Text == "1") currentSetName = (Convert.ToInt32(currentSetName) + 1).ToString();
                        if (currentSetName == textBox5.Text) currentSetName = "Roof";
                        
                        if(!db.status["CHiCHi"].floors.ContainsKey(currentSetName))
                            db.status["CHiCHi"].floors.Add(currentSetName, new dbFloorStatus());
                    }
                    else if (i1 == 1)
                    {
                        line = Regex.Replace(line, @"\s+", " ");
                        string[] dZx = line.Split(new char[1] { ' ' });
                        if (dZx.Length < 2) continue;
                        Dictionary<string, double> pss = new Dictionary<string, double>();
                        if (line.IndexOf("--") < 0 && line.IndexOf("Time") < 0)
                        {
                            //+一节点
                            //pss.Add("DX", Convert.ToDouble(dZx[2]));
                            //+一层
                            //ps.Add(countFloor, pss);
                            //countFloor++;
                            
                            //计算时程单位间隔时间
                            if (sec == -1) sec = Convert.ToDouble(dZx[1]);
                            else if (Global.sec == -1) Global.sec = Math.Round(Math.Abs(sec - Convert.ToDouble(dZx[1])),5);
                            //Global.splitTime = sec;
                            //记录最大计算时间
                            if (Global.MidasModelCalTime < Convert.ToDouble(dZx[1])) Global.MidasModelCalTime = Convert.ToDouble(dZx[1]);

                            if(!db.status["CHiCHi"].floors[currentSetName].timeStat.ContainsKey(Convert.ToDouble(dZx[1])))
                                db.status["CHiCHi"].floors[currentSetName].timeStat.Add(Convert.ToDouble(dZx[1]), new dbPointStatus());
                            if (currentSetDir == "X")
                            {
                                db.status["CHiCHi"].floors[currentSetName].timeStat[Convert.ToDouble(dZx[1])].fxG = Convert.ToDouble(dZx[2]);
                                db.status["CHiCHi"].floors[currentSetName].totalFX++;
                            }
                            else
                            {
                                db.status["CHiCHi"].floors[currentSetName].timeStat[Convert.ToDouble(dZx[1])].fyG = Convert.ToDouble(dZx[2]);
                                db.status["CHiCHi"].floors[currentSetName].totalFY++;
                            }
                        }
                        else if (dZx[1].IndexOf("--") > -1 && 
                            (
                            (db.status["CHiCHi"].floors[currentSetName].totalFX > 5 && currentSetDir=="X")
                            ||
                            (db.status["CHiCHi"].floors[currentSetName].totalFY > 5 && currentSetDir == "Y")
                            )
                            )//finish
                        {
                            //floorStaus.Add(currentSetName, ps);
                            //db.status["CHiCHi"].floors.Add(currentSetName, dbf);
                            i1 = 0;
                            MainForm.SafeInvoke(this.listBox1, new TaskDelegate(delegate()
                            {
                                int selIndex = listBox1.FindString(currentSetName);
                                if (selIndex < 0)
                                    selIndex = listBox1.Items.Add(currentSetName);
                                //int selIndex = listBox1.Items.Add(currentSetName+"| F"+ currentSetDir +" | 时程点集");
                                listBox1.SelectedIndex = (selIndex);
                                if (!debugMode) showDate(true, "F" + currentSetDir);
                            }));
                        }
                    }
                }
            }


            MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
            {
                listBox2.Items.Add("读取时程[CHiCHi]分析结果[U]");
            }));

            if (File.Exists(filePath + filesU[thisComboBox1Text]))
                sr = new StreamReader(filePath + filesU[thisComboBox1Text], Encoding.Default);
            if (sr != null)
            {
                i1 = 0;
                while ((line = sr.ReadLine()) != null)
                {
                    if (!debugMode)
                    MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                    {
                        int selIndex = listBox2.Items.Add(line.ToString());
                        if (i1 == 0 )
                            listBox2.SelectedIndex = (selIndex);
                    }));


                    if (i1 == 0 && line.IndexOf("**") > -1 && line.IndexOf("midas") < 0 && line.IndexOf("X-Axis") < 0)
                    {
                        countFloor = 1;
                        i1 = 1;
                        ps = new Dictionary<int, Dictionary<string, double>>();
                        string[] dddddd = Regex.Replace(line, @"\s+", " ").Split(new char[1] { ' ' });
                        //currentSetName = dddddd[3];

                        currentSetName = dddddd[3].Replace("aa", "").Replace("F", "").Replace("AA", "");//+dddddd[1]
                        currentSetDir = "X";
                        if (currentSetName == "wd") currentSetName = "Roof";
                        if (currentSetName == "屋顶") currentSetName = "Roof";
                        //if (currentSetName == textBox5.Text) currentSetName = "Roof";
                        if (!db.status["CHiCHi"].floors.ContainsKey(currentSetName))
                            db.status["CHiCHi"].floors.Add(currentSetName, new dbFloorStatus());
                    }
                    else if (i1 == 1)
                    {
                        line = Regex.Replace(line, @"\s+", " ");
                        string[] dZx = line.Split(new char[1] { ' ' });
                        if (dZx.Length < 2) continue;
                        Dictionary<string, double> pss = new Dictionary<string, double>();
                        if (line.IndexOf("--") < 0 && line.IndexOf("Axis") < 0)
                        {
                            //+一节点
                            //pss.Add("DX", Convert.ToDouble(dZx[2]));
                            //+一层
                            //ps.Add(countFloor, pss);
                            //countFloor++;

                            if (!db.status["CHiCHi"].floors[currentSetName].timeStat.ContainsKey(Convert.ToDouble(dZx[1])))
                                db.status["CHiCHi"].floors[currentSetName].timeStat.Add(Convert.ToDouble(dZx[1]), new dbPointStatus());
                            if (currentSetDir == "X")
                            {
                                db.status["CHiCHi"].floors[currentSetName].timeStat[Convert.ToDouble(dZx[1])].uxG = Convert.ToDouble(dZx[2]);
                                db.status["CHiCHi"].floors[currentSetName].totalUX++;
                            }
                            else
                            {
                                db.status["CHiCHi"].floors[currentSetName].timeStat[Convert.ToDouble(dZx[1])].uyG = Convert.ToDouble(dZx[2]);
                                db.status["CHiCHi"].floors[currentSetName].totalUY++;
                            }
                        }
                        else if (dZx[1].IndexOf("--") > -1  && 
                            (
                            (db.status["CHiCHi"].floors[currentSetName].totalUX > 5 && currentSetDir=="X")
                            ||
                            (db.status["CHiCHi"].floors[currentSetName].totalUY > 5 && currentSetDir == "Y")
                            )
                            )//finish
                        {
                            //floorStaus.Add(currentSetName, ps);
                            i1 = 0;
                            MainForm.SafeInvoke(this.listBox1, new TaskDelegate(delegate()
                            {
                                int selIndex=listBox1.FindString(currentSetName);
                                if (selIndex<0)
                                    selIndex = listBox1.Items.Add(currentSetName);
                                listBox1.SelectedIndex = (selIndex);
                                if (!debugMode) showDate(true, "U" + currentSetDir);
                            }));
                        }
                    }
                }
            }


            //***********************************


            MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
            {
                listBox2.Items.Add("读取时程[CHiCHi]分析结果[消能器单元txt结果]");
            }));

            if (File.Exists(filePath + filesU[thisComboBox1Text]))
                sr = new StreamReader(filePath + filesJ[thisComboBox1Text], Encoding.Default);
            if (sr != null)
            {
                i1 = 0;
                while ((line = sr.ReadLine()) != null)
                {
                    if (!debugMode)
                    MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                    {
                        int selIndex = listBox2.Items.Add(line.ToString());
                        if (i1 == 0)
                            listBox2.SelectedIndex = (selIndex);
                    }));


                    if (i1 == 0 && line.IndexOf("**") > -1 && line.IndexOf("midas") < 0 && line.IndexOf("X-Axis") < 0)
                    {
                        countFloor = 1;
                        i1 = 1;
                        ps = new Dictionary<int, Dictionary<string, double>>();
                        string[] dddddd = Regex.Replace(line, @"\s+", " ").Split(new char[1] { ' ' });
                        //currentSetName = dddddd[3];

                        currentSetName = dddddd[3].Replace("s", "").Replace("maxd", "").Replace("maxl", "");
                        currentSetDir = (
                            dddddd[3].IndexOf("s") > -1 ||
                            dddddd[3].IndexOf("maxd") > -1
                            ) ? "s" : "f";
                        if (!db.status["CHiCHi"].floors.ContainsKey(currentSetName))
                            db.status["CHiCHi"].floors.Add(currentSetName, new dbFloorStatus());
                    }
                    else if (i1 == 1)
                    {
                        line = Regex.Replace(line, @"\s+", " ");
                        string[] dZx = line.Split(new char[1] { ' ' });
                        if (dZx.Length < 2) continue;
                        //Dictionary<string, double> pss = new Dictionary<string, double>();
                        if (line.IndexOf("--") < 0 && line.IndexOf("Axis") < 0)
                        {
                            //+一节点
                            //pss.Add("DX", Convert.ToDouble(dZx[2]));
                            //+一层
                            //ps.Add(countFloor, pss);
                            //countFloor++;

                            if (!db.status["CHiCHi"].floors[currentSetName].timeStat.ContainsKey(Convert.ToDouble(dZx[1])))
                                db.status["CHiCHi"].floors[currentSetName].timeStat.Add(Convert.ToDouble(dZx[1]), new dbPointStatus());
                            if (currentSetDir == "s")
                            {
                                db.status["CHiCHi"].floors[currentSetName].timeStat[Convert.ToDouble(dZx[1])].ux = Convert.ToDouble(dZx[2]);
                                db.status["CHiCHi"].floors[currentSetName].totalUX++;
                            }
                            else
                            {
                                db.status["CHiCHi"].floors[currentSetName].timeStat[Convert.ToDouble(dZx[1])].fx = Convert.ToDouble(dZx[2]);
                                db.status["CHiCHi"].floors[currentSetName].totalFX++;
                            }
                            
                        }
                        else if (dZx[1].IndexOf("--") > -1 &&
                            (
                            (db.status["CHiCHi"].floors[currentSetName].totalUX > 5 && currentSetDir == "s")
                            ||
                            (db.status["CHiCHi"].floors[currentSetName].totalFX > 5 && currentSetDir == "f")
                            )
                            )//finish
                        {
                            //floorStaus.Add(currentSetName, ps);
                            i1 = 0;
                            db.status["CHiCHi"].floors[currentSetName].ifFloor = false;
                            //

                            //
                            MainForm.SafeInvoke(this.listBox1, new TaskDelegate(delegate()
                            {
                                int selIndex = listBox1.FindString(currentSetName);
                                if (selIndex < 0)
                                    selIndex = listBox1.Items.Add(currentSetName);
                                listBox1.SelectedIndex = (selIndex);
                                if (!debugMode) showDate(true, "滞回曲线");
                            }));
                        }
                    }
                }
            }


            //**************************导入excel塑性铰
            ExcelHelper ds;
            Excel.Worksheet ws;
            Array ar;
            //string curGK;
            foreach (string fileName in filesX)
            {
                currentGK = fileName.Replace(".xls", "");
                if (db.status["CHiCHi"].floors.ContainsKey(currentGK)) continue;
                if (File.Exists(filePath + "\\excel\\" + fileName))
                {
                    ds = new ExcelHelper(filePath + "\\excel\\" + fileName);
                    ws = ds.GetWorksheetByName("Time-History");
                    ar = ds.GetContent("Time-History");
                    //string[] sta = ar;
                    if(!db.status["CHiCHi"].floors.ContainsKey(currentGK))
                        db.status["CHiCHi"].floors.Add(currentGK, new dbFloorStatus());
                    db.status["CHiCHi"].floors[currentGK].ifFloor = false;
                    double time = 0.00;
                    for (int i = 1; i < 501; i++)
                    {
                        time = time + Global.sec;
                        time = Convert.ToDouble(time.ToString("F2"));

                        /*
                        MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                        {
                            listBox2.SelectedIndex =listBox2.Items.Add(time);
                        }));
                        */
                        db.status["CHiCHi"].floors[currentGK].timeStat.Add(time, new dbPointStatus());
                        db.status["CHiCHi"].floors[currentGK].timeStat[time].ux = 
                            Convert.ToDouble(ar.GetValue(i,1).ToString());
                        db.status["CHiCHi"].floors[currentGK].timeStat[time].fx =
                            Convert.ToDouble(ar.GetValue(i, 2).ToString());
                        db.status["CHiCHi"].floors[currentGK].totalFX++;
                        db.status["CHiCHi"].floors[currentGK].totalUX++;
                    }

                    //ar[1,10];
                    ds.Close();
                    //break;
                }
                MainForm.SafeInvoke(this.listBox1, new TaskDelegate(delegate()
                {
                    listBox1.SelectedIndex = listBox1.Items.Add(currentGK);
                    if (!debugMode) showDate(true, "滞回曲线");
                }));
                MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                {
                    listBox2.SelectedIndex =
                    listBox2.Items.Add("读取时程[CHiCHi]分析结果[消能器单元xls结果][" + fileName+"]");
                }));
                
            }


            //**************************
            #region 节点导入
            /*
            MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
            {
                listBox2.Items.Add("读取时程分析结果（节点）");
            }));

            if (File.Exists(filePath + "节点位移 9927.txt"))
                sr = new StreamReader(filePath + "节点位移 9927.txt", Encoding.Default);
            if (sr != null)
            {
                i1 = 0;
                while ((line = sr.ReadLine()) != null)
                {
                    MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                    {
                        listBox2.Items.Add(line.ToString());
                    }));


                    if (i1 == 0 && line.IndexOf("TIME") > -1 && line.IndexOf("DX") > -1)
                    {
                        countFloor = 1;
                        i1 = 1;
                        ps = new Dictionary<int, Dictionary<string, double>>();
                    }
                    if (i1 == 1)
                    {
                        line = Regex.Replace(line, @"\s+", " ");
                        string[] dZx = line.Split(new char[1] { ' ' });
                        if (dZx.Length < 2) break;
                        Dictionary<string, double> pss = new Dictionary<string, double>();
                        if (dZx[1].IndexOf("-")<0 && dZx[1]!="TIME")
                        {
                            //+一节点
                            pss.Add("Time", Convert.ToDouble(dZx[1]));
                            pss.Add("DX", Convert.ToDouble(dZx[2]));
                            pss.Add("DY", Convert.ToDouble(dZx[3]));
                            pss.Add("DZ", Convert.ToDouble(dZx[4]));
                            pss.Add("RX", Convert.ToDouble(dZx[5]));
                            pss.Add("RY", Convert.ToDouble(dZx[6]));
                            pss.Add("RZ", Convert.ToDouble(dZx[7]));
                            //+一层
                            ps.Add(countFloor, pss);
                            countFloor++;
                        }
                    }
                }
                floorStaus.Add(countGK.ToString(), ps);
                //db.status.Add(dbs);
                MainForm.SafeInvoke(this.listBox1, new TaskDelegate(delegate()
                {
                    int selIndex = listBox1.Items.Add(countGK.ToString() + "|  " + currentGK);
                    listBox1.SelectedIndex = (selIndex);
                    showDate(true);
                }));
            }
             */
            #endregion
            //**************************
            Dictionary<double, double> dbbbb = new Dictionary<double, double>();
            
            if (File.Exists(filePath + "ekx.txt"))
                sr = new StreamReader(filePath + "ekx.txt", Encoding.Default);
            if (sr != null)
            {
                i1 = 0;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line == "" || line=="	" || line == null) continue;

                    string text1 = line.Split(new char[1] { '	' })[0];
                    string text2 = line.Split(new char[1] { '	' })[1];
                    if (text1 == "" || text1 == null) continue;
                    dbbbb.Add(Convert.ToDouble(text1), Convert.ToDouble(text2));
                }
                MainForm.SafeInvoke(this.chart2, new TaskDelegate(delegate()
                {
                        foreach (KeyValuePair<double, double> kv in dbbbb)
                              this.chart2.Series[0].Points.AddXY(kv.Key, kv.Value);
                        this.chart2.ChartAreas[0].CursorX.Position = 20;
                }));
            }


            //midasResultFetch midasResult = new midasResultFetch(this.textBox1.Text);
            if (thread != null)
            {
                ifendsat update = new ifendsat(ifend);
                this.Invoke(update, true);  //调用窗体Invoke方法  
                thread.Abort();
            }
        }

        private void cal()
        {
            //double Ws=0;
            MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
            {
                listBox2.SelectedIndex = listBox2.Items.Add("准备计算数据");
            }));

            double Wsi = 0;

            double fv = Convert.ToDouble(this.dataGridView2.Rows[0].Cells[1].Value);//235
            double s = Convert.ToDouble(this.dataGridView2.Rows[1].Cells[1].Value);//50
            double l = Convert.ToDouble(this.dataGridView2.Rows[2].Cells[1].Value);//6.5
            double a = Convert.ToDouble(this.dataGridView2.Rows[3].Cells[1].Value) / 180 * Math.PI;//45

            MainForm.SafeInvoke(this.listBox1, new TaskDelegate(delegate()
            {
                int selIndex = listBox1.FindString("楼层Total");
                if (selIndex < 0)
                    selIndex = listBox1.Items.Add("楼层Total");
                selIndex = listBox1.FindString("消能器Total");
                if (selIndex < 0)
                    selIndex = listBox1.Items.Add("消能器Total");

                listBox1.SelectedIndex = (selIndex);
                if (!db.status["CHiCHi"].floors.ContainsKey("楼层Total"))
                    db.status["CHiCHi"].floors.Add("楼层Total", new dbFloorStatus());
                else db.status["CHiCHi"].floors["楼层Total"] = new dbFloorStatus();
                if (!db.status["CHiCHi"].floors.ContainsKey("消能器Total"))
                    db.status["CHiCHi"].floors.Add("消能器Total", new dbFloorStatus());
                else db.status["CHiCHi"].floors["消能器Total"] = new dbFloorStatus();
            }));

            foreach (KeyValuePair<string, dbFloorStatus> gks in db.status["CHiCHi"].floors)
            {

                if (gks.Key.IndexOf("Total") > -1) continue;
                if (!gks.Value.ifFloor)
                {//每个消能器
                    //Wsi = 0;

                    //搜寻滞回曲线弹性阶段
                    MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                    {
                        listBox2.SelectedIndex = listBox2.Items.Add("搜寻滞回曲线[" + gks.Key + "]弹性阶段");
                    }));
                    KSet ks = new KSet();
                    foreach (KeyValuePair<double, dbPointStatus> gksTime in gks.Value.timeStat)
                    {//滞回曲线弹性阶段平均K值
                        double averageK = 0;
                        double nextSecName = (Convert.ToDouble(gksTime.Key) + Global.sec);
                        if (gks.Value.timeStat.ContainsKey(nextSecName))
                            if ((gksTime.Value.ux - gks.Value.timeStat[nextSecName].ux) > 0)
                                averageK = Math.Abs((gksTime.Value.fx - gks.Value.timeStat[nextSecName].fx) / (gksTime.Value.ux - gks.Value.timeStat[nextSecName].ux));

                        if (db.status["CHiCHi"].floors[gks.Key].k == 0)
                            db.status["CHiCHi"].floors[gks.Key].k = averageK;
                        if (
                            Math.Abs(averageK - db.status["CHiCHi"].floors[gks.Key].k) / Math.Max(averageK, db.status["CHiCHi"].floors[gks.Key].k) < 0.1
                            )
                            db.status["CHiCHi"].floors[gks.Key].k = (averageK + db.status["CHiCHi"].floors[gks.Key].k) / 2;

                        if (db.status["CHiCHi"].floors[gks.Key].maxy < gksTime.Value.fx)
                            db.status["CHiCHi"].floors[gks.Key].maxy = gksTime.Value.fx;
                        if (db.status["CHiCHi"].floors[gks.Key].miny > gksTime.Value.fx)
                            db.status["CHiCHi"].floors[gks.Key].miny = gksTime.Value.fx;
                        //保存k瞬时值到数据库，统计得到点
                        ks.add(averageK.ToString("F1"));//保留一位小数
                    }
                    db.status["CHiCHi"].floors[gks.Key].k1 = Math.Min(Convert.ToDouble(ks.get(1)), Convert.ToDouble(ks.get(2)));
                    //按全部瞬时值的前两个众数，保存k1值。


                    double tanxingD = Math.Max(Math.Abs(db.status["CHiCHi"].floors[gks.Key].maxy), Math.Abs(db.status["CHiCHi"].floors[gks.Key].miny)) / db.status["CHiCHi"].floors[gks.Key].k; //根据滞回曲线求的min拉或压的最大弹性变形值
                    //tanxingD = db.status["CHiCHi"].floors[gks.Key].maxy / db.status["CHiCHi"].floors[gks.Key].k;

                    db.status["CHiCHi"].floors[gks.Key].tanxingD = tanxingD;

                    //用于求的弹性阶段四个点坐标，显示图表
                    /*
                    db.status["CHiCHi"].floors[gks.Key].tanxinD1.ux = db.status["CHiCHi"].floors[gks.Key].maxy / db.status["CHiCHi"].floors[gks.Key].k;
                    db.status["CHiCHi"].floors[gks.Key].tanxinD1.fx = db.status["CHiCHi"].floors[gks.Key].maxy;
                    db.status["CHiCHi"].floors[gks.Key].tanxinD2.ux = db.status["CHiCHi"].floors[gks.Key].miny / db.status["CHiCHi"].floors[gks.Key].k;
                    db.status["CHiCHi"].floors[gks.Key].tanxinD2.fx = db.status["CHiCHi"].floors[gks.Key].miny;
                    */
                    //计算Ws
                    MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                    {
                        listBox2.SelectedIndex = listBox2.Items.Add("计算消能器总耗能[叠加" + gks.Key + "]");
                    }));
                    foreach (KeyValuePair<double, dbPointStatus> gksTime in gks.Value.timeStat)
                    {
                        if (gksTime.Value.fx > gks.Value.maxFx) gks.Value.maxFx = gksTime.Value.fx;
                        if (gksTime.Value.fx < gks.Value.minFx) gks.Value.minFx = gksTime.Value.fx;
                    }
                    foreach (KeyValuePair<double, dbPointStatus> gksTime in gks.Value.timeStat)
                    {
                        double newKey = gksTime.Key;
                        db.status["CHiCHi"].floors["消能器Total"].pAdd(newKey, new dbPointStatus());

                        double tempWs = Math.Abs(4 * fv * 0.001 * meterial2Area[elementId2meterial[gks.Key]] * cutTanxingjieduan(gksTime.Value.ux, tanxingD));
                        //tempWs = Math.Abs(4 * gksTime.Value.fx * cutTanxingjieduan(gksTime.Value.ux, tanxingD));

                        db.status["CHiCHi"].floors["消能器Total"].timeStat[newKey].ws += tempWs;
                        db.status["CHiCHi"].floors[gks.Key].timeStat[newKey].ws = tempWs;
                    }
                }
                else//每层
                {
                    Wsi = 0;
                    MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                    {
                        listBox2.SelectedIndex = listBox2.Items.Add("计算楼层总耗能[叠加" + gks.Key + "]");
                    }));

                    //用于包络计算方法，不对各时刻积分，只取最大值。
                    double currentFloorMaxFi = 0;
                    double currentFloorMaxFiTime = 0;
                    double currentFloorMaxFiG = 0;
                    double currentFloorMaxFiGTime = 0;
                    double currentFloorMaxUi = 0;
                    double currentFloorMaxUiTime = 0;
                    double currentFloorMinFi = 0;
                    double currentFloorMinFiTime = 0;
                    double currentFloorMinUi = 0;
                    double currentFloorMinUiTime = 0;
                    double currentFloorMaxWs = 0;
                    double currentFloorMaxWsTime = 0;
                    double currentFloorMinWs = 0;
                    double currentFloorMinWsTime = 0;

                    foreach (KeyValuePair<double, dbPointStatus> gksTime in gks.Value.timeStat)//每个时刻
                    {
                        //处理"层间"xx

                        double topF = 0, btmF = 0;
                        string topFloorNum = "", btmFloorNum = "";

                        if (gks.Key == "Roof")
                        {
                            topFloorNum = "Roof";
                            btmFloorNum = (Convert.ToInt32(textBox5.Text) - 1).ToString();
                        }
                        else if (gks.Key == (Convert.ToInt32(textBox5.Text) - 1).ToString())
                        {
                            topFloorNum = "Roof";
                            btmFloorNum = (Convert.ToInt32(gks.Key) - 1).ToString();
                        }
                        else if (gks.Key == "1")
                        {
                            topFloorNum = (Convert.ToInt32(gks.Key) + 1).ToString();
                            btmFloorNum = "1";
                        }
                        else
                        {
                            topFloorNum = (Convert.ToInt32(gks.Key) + 1).ToString();
                            btmFloorNum = (Convert.ToInt32(gks.Key) - 1).ToString();
                        }

                        //if (gks.Key == "Roof") { topF = 0; }

                        if (gks.Key == "Roof") topF = 0;
                        else if (db.status["CHiCHi"].floors.ContainsKey(topFloorNum.ToString()))
                        {
                            topF = db.status["CHiCHi"].floors[topFloorNum.ToString()].timeStat[gksTime.Key].fxG;
                            //btmU = db.status["CHiCHi"].floors[btmFloorNum.ToString()].timeStat[gksTime.Key].uxG;
                        }

                        gks.Value.timeStat[gksTime.Key].fx = (gks.Value.timeStat[gksTime.Key].fxG) - (topF);
                        gks.Value.timeStat[gksTime.Key].ux = gks.Value.timeStat[gksTime.Key].uxG;


                        db.status["CHiCHi"].floors["楼层Total"].pAdd(gksTime.Key, new dbPointStatus());
                        db.status["CHiCHi"].floors["楼层Total"].timeStat[gksTime.Key].ws +=
                            (0.5 * (gks.Value.timeStat[gksTime.Key].fxG - topF) * gksTime.Value.ux);
                        //Ws = 1/2 * （Fv层间）* ui
                        gksTime.Value.ws =
                            (0.5 * (gks.Value.timeStat[gksTime.Key].fxG - topF) * gksTime.Value.ux);//瞬时值



                        //用于包络计算方法，求最大值。
                        if ((gksTime.Value.fx) > (currentFloorMaxFi))
                        {
                            currentFloorMaxFi = (gksTime.Value.fx);
                            currentFloorMaxFiTime = gksTime.Key;
                        }

                        if ((gksTime.Value.ux) > (currentFloorMaxUi))
                        {
                            currentFloorMaxUi = (gksTime.Value.ux);
                            currentFloorMaxUiTime = gksTime.Key;
                        }
                        if ((gksTime.Value.fx) < (currentFloorMinFi))
                        {
                            currentFloorMinFi = (gksTime.Value.fx);
                            currentFloorMinFiTime = gksTime.Key;
                        }
                        if ((gksTime.Value.ux) < (currentFloorMinUi))
                        {
                            currentFloorMinUi = (gksTime.Value.ux);
                            currentFloorMinUiTime = gksTime.Key;
                        }
                        if ((gksTime.Value.ws) > currentFloorMaxWs)
                        {
                            currentFloorMaxWs = (gksTime.Value.ws);
                            currentFloorMaxWsTime = gksTime.Key;
                        }
                        if ((gksTime.Value.ws) < currentFloorMinWs)
                        {
                            currentFloorMinWs = (gksTime.Value.ws);
                            currentFloorMinWsTime = gksTime.Key;
                        }
                        //Ground
                        if (Math.Abs(gksTime.Value.fxG) > (currentFloorMaxFiG))
                        {
                            currentFloorMaxFiG = Math.Abs(gksTime.Value.fxG);
                            currentFloorMaxFiGTime = gksTime.Key;
                        }
                    }
                    gks.Value.maxFi = currentFloorMaxFi;
                    gks.Value.maxFiT = currentFloorMaxFiTime;
                    gks.Value.maxUi = currentFloorMaxUi;
                    gks.Value.maxUiT = currentFloorMaxUiTime;

                    gks.Value.minFi = currentFloorMinFi;
                    gks.Value.minFiT = currentFloorMinFiTime;
                    gks.Value.minUi = currentFloorMinUi;
                    gks.Value.minUiT = currentFloorMinUiTime;

                    gks.Value.maxWs = currentFloorMaxWs;
                    gks.Value.maxWsT = currentFloorMaxWsTime;
                    gks.Value.minWs = currentFloorMinWs;
                    gks.Value.minWsT = currentFloorMinWsTime;

                    gks.Value.maxFiG = currentFloorMaxFiG;
                    gks.Value.maxFiGT = currentFloorMaxFiGTime;

                }
                //计算完成
            }
            MainForm.SafeInvoke(this.listBox1, new TaskDelegate(delegate()
            {
                listBox1.SelectedIndex = listBox1.Items.IndexOf("楼层Total");
                showDate(true, "Ws");
                listBox1.SelectedIndex = listBox1.Items.IndexOf("消能器Total");
                showDate(true, "Ws");
                //Thread.Sleep(1000);
                showDate(true, "znb");
                //Thread.Sleep(1000);
                showDate(true, "znb25");
            }));

            //包络计算方法：
            ////遍历读取楼层最大ws的时刻点
            double maxTime = 0;
            double maxWs = 0;
            foreach (KeyValuePair<double, dbPointStatus> gksTime in db.status["CHiCHi"].floors["楼层Total"].timeStat)
            {
                if (gksTime.Value.ws > maxWs)
                {
                    maxWs = gksTime.Value.ws;
                    maxTime = gksTime.Key;
                    Global.maxtime = maxTime;
                }
            }
            MainForm.SafeInvoke(this.label13, new TaskDelegate(delegate()
            {
                this.label13.Text = "楼层Ws最大值时刻：" + maxTime + "s (Ws=" + maxWs.ToString("E2") + ")";
            }));

            //MessageBox.Show("Max Ws Time=" + maxTime.ToString());//获得maxTime的计算范围
            if (checkBox4.Checked)
                maxTime = Global.MidasModelCalTime;
            else maxTime = maxTime + Convert.ToDouble(textBox8.Text);

            double xnqWsDuringSpecPred = 0;
            double xnqTotalHN = 0;
            foreach (KeyValuePair<string, dbFloorStatus> gks in db.status["CHiCHi"].floors)
            {
                if (gks.Key.IndexOf("Total") > -1) continue;
                if (!gks.Value.ifFloor)
                {//每个消能器
                    //效能器时程无效时间段清除（另外封装的函数，从末端开始，把等斜率的那段去除）
                    double maxBlTime = dbMathPoint.zhqxEdge(gks.Value.timeStat, maxTime);
                    //double zhqxArea = 0;
                    double currentXnqK = 0;

                    //获取滞回曲线ux中点
                    double maxUx = 0;
                    double minUx = 0;
                    double maxFx = 0;
                    double minFx = 0;
                    foreach (KeyValuePair<double, dbPointStatus> gksTime in gks.Value.timeStat)
                    {
                        if (gksTime.Key > maxBlTime) break;//超出段停止
                        gksTime.Value.ux = gksTime.Value.ux - gksTime.Value.fx / gks.Value.k;//拉伸便于数据整理
                        gksTime.Value.fx = gksTime.Value.fx - gksTime.Value.ux * gks.Value.k1;
                        if (gksTime.Value.ux > maxUx) maxUx = gksTime.Value.ux;
                        else if (gksTime.Value.ux < minUx) minUx = gksTime.Value.ux;
                        if (gksTime.Value.fx > maxFx) maxFx = gksTime.Value.fx;
                        else if (gksTime.Value.fx < minFx) minFx = gksTime.Value.fx;
                    }
                    foreach (KeyValuePair<double, dbPointStatus> gksTime in gks.Value.timeStat)
                    {
                        if (gksTime.Key > maxBlTime) break;//超出段停止
                        gksTime.Value.fx = gksTime.Value.fx + gksTime.Value.ux * gks.Value.k1;
                        gksTime.Value.ux = gksTime.Value.ux + gksTime.Value.fx / gks.Value.k;//取消拉伸 恢复原来
                    }
                    //滞回曲线外包面积
                    //gks.Value.zhqxArea = Math.Abs((maxUx - minUx) * (maxFx - minFx));

                    gks.Value.d0 = new dbPointStatus(maxUx + maxFx / gks.Value.k, maxFx + maxUx * gks.Value.k1, "x");
                    gks.Value.d1 = new dbPointStatus(minUx + maxFx / gks.Value.k, maxFx + minUx * gks.Value.k1, "x");
                    gks.Value.d2 = new dbPointStatus(minUx + minFx / gks.Value.k, minFx + minUx * gks.Value.k1, "x");
                    gks.Value.d3 = new dbPointStatus(maxUx + minFx / gks.Value.k, minFx + maxUx * gks.Value.k1, "x");
                    //MessageBox.Show(gks.Key + ":" + gks.Value.zhqxArea);
                    gks.Value.zhqxArea = dbMathPoint.sBy3P(gks.Value.d0, gks.Value.d1, gks.Value.d3);


                    MainForm.SafeInvoke(this.listBox2, new TaskDelegate(delegate()
                    {
                        listBox2.SelectedIndex = listBox2.Items.Add("滞回曲线凸包算法计算: #" + gks.Key);
                    }));

                    if (gks.Value.zhqxArea > 30)
                    {
                        Form1 fff = new Form1();
                        var nodes = new List<PointF>();
                        foreach (KeyValuePair<double, dbPointStatus> gksTime in gks.Value.timeStat)
                        {
                            if (gksTime.Key > maxBlTime) break;//超出段停止
                            nodes.Add(new PointF((float)Math.Round(gksTime.Value.ux, 3), (float)Math.Round(gksTime.Value.fx, 0)));
                        }
                        fff.zhqxNodes = nodes;
                        fff.runProgram();
                        gks.Value.zhqxArea = fff.area;
                        gks.Value.zhqxAreaFxString = fff.fxSTR;
                        gks.Value.zhqxAreaUxString = fff.uxSTR;
                    }


                    xnqTotalHN += gks.Value.zhqxArea;
                    //MessageBox.Show(gks.Key + ":" + gks.Value.zhqxArea);
                }
            }

            //包络法计算楼层应变能
            double floorHN = 0;
            foreach (KeyValuePair<string, dbFloorStatus> gks in db.status["CHiCHi"].floors)
            {
                if (gks.Key.IndexOf("Total") > -1) continue;
                if (gks.Value.ifFloor)
                {//楼层

                    double topFG = 0;
                    string topFloorNum = "";
                    if (gks.Key == "Roof" || gks.Key == (Convert.ToInt32(textBox5.Text)).ToString())
                        topFloorNum = "Roof";
                    else if (gks.Key == (Convert.ToInt32(textBox5.Text) - 1).ToString())
                        topFloorNum = "Roof";
                    else
                        topFloorNum = (Convert.ToInt32(gks.Key) + 1).ToString();

                    double u = Math.Max(db.status["CHiCHi"].floors[gks.Key].maxUi, -db.status["CHiCHi"].floors[gks.Key].minUi);
                    double fG = db.status["CHiCHi"].floors[gks.Key].maxFiG;
                    double fGtop = db.status["CHiCHi"].floors[topFloorNum].maxFiG;
                    db.status["CHiCHi"].floors[gks.Key].maxFiG_CutTopFloor_maxFiG = (fG - fGtop);
                    floorHN += 0.5 * (fG - fGtop) * u;
                    
                    double fcj = db.status["CHiCHi"].floors[gks.Key].maxFi;
                    if(db.status["CHiCHi"].floors[gks.Key].maxFi<-db.status["CHiCHi"].floors[gks.Key].minFi)
                        fcj = db.status["CHiCHi"].floors[gks.Key].minFi;
                    //floorHN += 0.5 * fcj * u;
                    


                }
            }
            floorHN = Math.Abs(floorHN);

            MainForm.SafeInvoke(this.label16, new TaskDelegate(delegate()
            {
                label16.Text = "附加有效阻尼比包络值 = " + xnqTotalHN.ToString("E2") + " / 4 pi " + floorHN.ToString("E2");
            }));
            Global.resultZnbBl = (xnqTotalHN * 100 / (4 * 3.1415 * floorHN));
            MainForm.SafeInvoke(this.label17, new TaskDelegate(delegate()
            {
                label17.Text = "= " + Global.resultZnbBl.ToString("F2") + "%";
            }));


            //综合法：
            double zhMaxWs = 0;
            foreach (KeyValuePair<double, dbPointStatus> pDb in db.status["CHiCHi"].floors["楼层Total"].timeStat)
            {
                if (zhMaxWs < pDb.Value.ws) zhMaxWs = pDb.Value.ws;
            }
            MainForm.SafeInvoke(this.label19, new TaskDelegate(delegate()
            {
                label19.Text = "附加阻尼比 = " + xnqTotalHN.ToString("E2") + "/ 4 pi " + zhMaxWs.ToString("E2") + "=" + (xnqTotalHN / (4 * 3.1415 * zhMaxWs) * 100).ToString("F2") + "%";
            }));
            Global.resultZnbZh = (xnqTotalHN / (4 * 3.1415 * zhMaxWs) * 100);

            MainForm.SafeInvoke(this.button3, new TaskDelegate(delegate()
            {
                button3.Enabled = true;
            }));
            MainForm.SafeInvoke(this.chart3, new TaskDelegate(delegate()
            {
                button11_Click(null, null);
            }));
                        //dataGridView1.Rows.
            MainForm.SafeInvoke(this.dataGridView1, new TaskDelegate(delegate()
            {
                //dataGridView1.Rows[Convert.ToInt32(maxTime / Global.sec)-1].Selected = true;//设置为选中. 
                //dataGridView1.FirstDisplayedScrollingRowIndex = Convert.ToInt32(maxTime / Global.sec)-1;  //设置第一行显示
            }));

        }

        private void button6_Click(object sender, EventArgs e)
        {
            showDate();
        }

        private void toolStripStatusLabel2_Click(object sender, EventArgs e)
        {
            info info = new info();
            info.StartPosition = FormStartPosition.CenterParent;
            info.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            chart1.Series.Clear();
            //chart1.Series.Add(

        }

        public double cutTanxingjieduan(double d1t,double dv){
            if( Math.Abs( d1t )-dv<0)return 0;
            else return Math.Abs(d1t) - dv;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DateTime now=DateTime.Now;
            if (now.Year >= 2017 && now.Month >= 6) { 
                MessageBox.Show("软件已有更新，请点击界面右下角蓝色软件说明链接，联系开发商。"); 
                return; 
            }

            tabControl1.SelectedTab = tabControl1.TabPages[2];
            radioButton7.Enabled = true;
            radioButton8.Enabled = true;
            checkBox4.Enabled = false;
            textBox8.Enabled = false;
            button2.Enabled = false;
            button4.Enabled = false;
            thread = new Thread(new ThreadStart(cal));
            thread.Start();
            //cal();
        }



        private void button3_Click(object sender, EventArgs e)
        {
            button3.Enabled = false;
            thread = new Thread(new ThreadStart(expToExcel));
            thread.Start();

        }
        private void expToExcel()
        {
            creatExcel ce = new creatExcel();
            ce.st = Convert.ToDouble(textBox2.Text);
            ce.ed = Convert.ToDouble(textBox3.Text);
            ce.meterial2Area = meterial2Area;
            ce.elementId2meterial = elementId2meterial;
            ce.creat(filePath, db,this);

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex >=1) listBox1.SelectedIndex -= 1;
            if(checkBox2.Checked)
                chart1.Series.Clear();
            showDate();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex < listBox1.Items.Count -1 ) listBox1.SelectedIndex += 1;
            if (checkBox2.Checked)
                chart1.Series.Clear();
            showDate();
        }

        private void textBox4_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) {
                if (listBox1.Items.Contains(textBox4.Text))
                {
                    listBox1.SelectedIndex = listBox1.Items.IndexOf(textBox4.Text);
                }
            }
        }

        private void chart2_Click(object sender, EventArgs e)
        {
            chart1.ChartAreas[0].CursorX.Position = chart2.ChartAreas[0].CursorX.Position;
            chart1.ChartAreas[0].CursorX.SelectionStart = chart2.ChartAreas[0].CursorX.SelectionStart;
            chart1.ChartAreas[0].CursorX.SelectionEnd = chart2.ChartAreas[0].CursorX.SelectionEnd;
            textBox2.Text = chart1.ChartAreas[0].CursorX.SelectionStart.ToString();
            textBox3.Text = chart1.ChartAreas[0].CursorX.SelectionEnd.ToString();
        }

        private void chart1_Click(object sender, EventArgs e)
        {
            chart2.ChartAreas[0].CursorX.Position = chart1.ChartAreas[0].CursorX.Position;
            chart2.ChartAreas[0].CursorX.SelectionStart = chart1.ChartAreas[0].CursorX.SelectionStart;
            chart2.ChartAreas[0].CursorX.SelectionEnd = chart1.ChartAreas[0].CursorX.SelectionEnd;
            textBox2.Text = chart1.ChartAreas[0].CursorX.SelectionStart.ToString();
            textBox3.Text = chart1.ChartAreas[0].CursorX.SelectionEnd.ToString();
            Global.currentSetTime = Math.Round(chart1.ChartAreas[0].CursorX.Position,2);
            button12_Click(sender,e);
            label11.Text = "时刻点：" + chart1.ChartAreas[0].CursorX.Position.ToString(); 
        }

        private void button9_Click(object sender, EventArgs e)
        {
            toolsGenMgt mgtForm = new toolsGenMgt();
            mgtForm.roofNum = Convert.ToInt32( textBox5.Text);
            mgtForm.textBox1.Text = filePath + "截面信息.xlsx";
            mgtForm.textBox2.Text = filePath + "单元信息.xlsx";
            mgtForm.textBox8.Text = filePath + "楼层单元信息.xlsx";
            //mgtForm.ShowDialog();
            if (mgtForm.ShowDialog() == DialogResult.OK)
            {
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            Form1 fff = new Form1();
            var nodes = new List<PointF>();
            foreach(KeyValuePair<double, dbPointStatus> gksTime in db.status["CHiCHi"].floors[textBox7.Text].timeStat){
                //nodes.Add(new PointF((float)(gksTime.Value.ux*20+300),(float)(gksTime.Value.fx/30+160)));
                nodes.Add(new PointF((float)Math.Round(gksTime.Value.ux,2)*100, (float)Math.Round(gksTime.Value.fx, 0)));
            }
            fff.zhqxNodes = nodes;
            if (fff.ShowDialog() == DialogResult.OK)
            {
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            filePath = @"E:\BRB\结果\江门\EL等全部结果\";
            filePathMethor(filePath);
            textBox5.Text = "29";
            textBox6.Text = "0";
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            filePath = @"E:\BRB\结果\横琴\横琴全部结果\";
            filePathMethor(filePath);
            textBox5.Text = "19";
            textBox6.Text = "0";
            textBox8.Text = "0";
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            filePath = @"E:\BRB\结果\新疆喀什\";
            filePathMethor(filePath);
            textBox5.Text = "4";
            textBox6.Text = "1";
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox4.Text = textBox7.Text = listBox1.Text;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if(checkBox5.Checked)
                chart3.Series.Clear();
            if (chart3.Series.IsUniqueName(comboBox2.Text))
                chart3.Series.Add(comboBox2.Text);

            this.dataGridView1.Rows.Clear();

            Series ser = chart3.Series.FindByName(comboBox2.Text);
            ser.Points.Clear();
            ser.ChartType = SeriesChartType.Line;

            double max = 0;
            foreach (KeyValuePair<string, dbFloorStatus> pDb in db.status["CHiCHi"].floors)
            {
                if (!pDb.Value.ifFloor || pDb.Key.ToString().IndexOf("Total") > -1) continue;
                int floorNum;
                if (pDb.Key == "Roof") floorNum = Convert.ToInt32(textBox5.Text);
                else floorNum = Convert.ToInt32(pDb.Key);
                double currentData=0;
                if(comboBox2.Text =="最大层剪力")
                    currentData=pDb.Value.maxFiG;
                    //ser.Points.AddXY(pDb.Value.maxFiG, floorNum);
                else if (comboBox2.Text == "最大层间剪力(各层时程的最大值，再层间相减)")
                    currentData=pDb.Value.maxFiG_CutTopFloor_maxFiG;
                    //ser.Points.AddXY(pDb.Value.maxFiG_CutTopFloor_maxFiG, floorNum);
                else if (comboBox2.Text == "最大层间剪力(层间相减，再求时程最大值)")
                    currentData=pDb.Value.maxFi;

                currentData = Math.Round(currentData);
                ser.Points.AddXY(currentData, floorNum);

                //添加表格
                int drId = this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[drId].Cells[0].Value = pDb.Key;
                this.dataGridView1.Rows[drId].Cells[1].Value = currentData;

                if (max < currentData) max = currentData;
            }


            if (Global.currentSetTime != 0)
            {
                if (chart3.Series.IsUniqueName("瞬时值"))
                    chart3.Series.Add("瞬时值");
                ser = chart3.Series.FindByName("瞬时值");
                ser.Points.Clear();
                ser.ChartType = SeriesChartType.Line;
                ///////
                foreach (KeyValuePair<string, dbFloorStatus> pDb in db.status["CHiCHi"].floors)
                {
                    if (!pDb.Value.ifFloor || pDb.Key.ToString().IndexOf("Total") > -1) continue;
                    int floorNum;
                    if (pDb.Key == "Roof") floorNum = Convert.ToInt32(textBox5.Text);
                    else floorNum = Convert.ToInt32(pDb.Key);
                    double currentData = 0;

                    if (!pDb.Value.timeStat.ContainsKey(Global.currentSetTime))
                    {
                        Global.currentSetTime = Convert.ToInt32(Global.currentSetTime / Global.sec) * Global.sec;
                    }
                    Global.currentSetTime = Math.Round(Global.currentSetTime, 2);

                    if(pDb.Value.timeStat[Global.currentSetTime].fxG!=0)
                        currentData = pDb.Value.timeStat[Global.currentSetTime].fx;
                    currentData = Math.Round( Math.Abs(currentData));
                    ser.Points.AddXY(currentData, floorNum);

                    if (max < currentData) max = currentData;
                }

            }
            //chart3.ChartAreas[0].AxisX.Crossing = 0;
            //chart3.ChartAreas[0].AxisX.MajorGrid.Interval = Math.Round(max/5000)*1000;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (Global.currentSetTime >= Global.MidasModelCalTime)
            {
                timer.Elapsed -= new System.Timers.ElapsedEventHandler(Timer1_Elapsed);  
                return;
            }
            Global.currentSetTime += Global.sec;
            MainForm.SafeInvoke(this.chart2, new TaskDelegate(delegate()
            {
               chart2.ChartAreas[0].CursorX.Position = chart1.ChartAreas[0].CursorX.Position = Global.currentSetTime;
            }));

            MainForm.SafeInvoke(this.chart3, new TaskDelegate(delegate()
            {
                if (Global.currentSetTime != 0)
                {
                    if (chart3.Series.IsUniqueName("瞬时值"))
                        chart3.Series.Add("瞬时值");
                    Series ser3 = chart3.Series.FindByName("瞬时值");
                    ser3.Points.Clear();
                    ser3.ChartType = SeriesChartType.Line;

                    if (chart4.Series.IsUniqueName("瞬时值"))
                        chart4.Series.Add("瞬时值");
                    Series ser4 = chart4.Series.FindByName("瞬时值");
                    ser4.Points.Clear();
                    ser4.ChartType = SeriesChartType.Line;
                    chart4.ChartAreas[0].AxisX.Maximum = 300;
                    chart4.ChartAreas[0].AxisX.Minimum = -300;
                    ///////
                    foreach (KeyValuePair<string, dbFloorStatus> pDb in db.status["CHiCHi"].floors)
                    {
                        if (!pDb.Value.ifFloor || pDb.Key.ToString().IndexOf("Total") > -1) continue;
                        int floorNum;
                        if (pDb.Key == "Roof") floorNum = Convert.ToInt32(textBox5.Text);
                        else floorNum = Convert.ToInt32(pDb.Key);
                        double currentData = 0;

                        if (!pDb.Value.timeStat.ContainsKey(Global.currentSetTime))
                            Global.currentSetTime = Convert.ToInt32(Global.currentSetTime / Global.sec) * Global.sec;
                        Global.currentSetTime = Math.Round(Global.currentSetTime, 2);


                        currentData = pDb.Value.timeStat[Global.currentSetTime].fx;
                        currentData = Math.Round(Math.Abs(currentData));
                        ser3.Points.AddXY(currentData, floorNum);

                        currentData = pDb.Value.timeStat[Global.currentSetTime].ux;
                        ser4.Points.AddXY(currentData, floorNum);

                    }
                }
            }));

        }
        System.Timers.Timer timer = new System.Timers.Timer();
        private void button13_Click(object sender, EventArgs e)
        {
            timer.Enabled = true;
            timer.Interval = 20;//执行间隔时间,单位为毫秒  
            timer.Start();
            timer.Elapsed += new System.Timers.ElapsedEventHandler(Timer1_Elapsed);  
            
        }
        private void Timer1_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            button12_Click(sender, e);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            timer.Elapsed -= new System.Timers.ElapsedEventHandler(Timer1_Elapsed);  
        }   
	}
}
