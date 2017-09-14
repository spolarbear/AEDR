using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;


using Aspose.Cells; //已经添加了应用 
using Aspose.Cells.Drawing;


namespace sunproject
{

    public class creatExcel
    {
        public Dictionary<int, double> meterial2Area = new Dictionary<int, double>();
        public Dictionary<string, int> elementId2meterial = new Dictionary<string, int>();
        public double st, ed;
        private string l2n(int n1,int n2){
            //string str = "";
            //CellsHelper.ConvertR1C1FormulaToA1(str,n1, n2);
            //return str;
            if (n2 >= 26)
            {
                char zm = (char)((n2 % 26) + 65);
                char zm10 = (char)((int)(n2 / 26) + 64);
                if (n2 == 26) return "AA"  + (n1 + 1).ToString();
                return zm10.ToString() + zm.ToString() + (n1 + 1).ToString();
            }
            else
            {
                char zm = (char)(n2 + 65);
                return zm.ToString() + (n1 + 1).ToString();
            }
        }
        private Worksheet getSheetByName(Workbook workbook ,string name){
            foreach (Worksheet sheetName in workbook.Worksheets)
            {
                if (sheetName.Name == name)
                {
                    return workbook.Worksheets[name];
                }
            }
            return workbook.Worksheets.Add(name);
        }
        private void initSheet(Worksheet sheet){
            sheet.Cells["A1"].PutValue("BRB附加有效阻尼比计算实用程序数据及计算过程Excel表 - 广东省建筑设计研究院ADG机场所");
            sheet.Cells["A2"].PutValue("楼层(消能器)编号");
            sheet.Cells["A12"].PutValue("时刻");
        }
        public void creat(string path,db db,MainForm form)
        {
            int chartMaxRow = Convert.ToInt32(Global.MidasModelCalTime / Global.sec) + 12;
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
            Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];//getSheetByName(workbook,"总信息");
            sheet.Name = "总信息";
            workbook.Worksheets.Add("楼层");
            //workbook.Worksheets.Add("消能器1");
            sheet.FreezePanes(1, 1, 1, 0);//冻结第一行

            workbook.Worksheets["总信息"].Cells["A3"].PutValue("时变法");
            workbook.Worksheets["总信息"].Cells["A4"].PutValue(Global.resultZnbSb.ToString("F2")+"");
            workbook.Worksheets["总信息"].Cells["A5"].PutValue("包络法");
            workbook.Worksheets["总信息"].Cells["A6"].PutValue(Global.resultZnbBl.ToString("F2") + "");
            workbook.Worksheets["总信息"].Cells["A7"].PutValue("综合法");
            workbook.Worksheets["总信息"].Cells["A8"].PutValue(Global.resultZnbZh.ToString("F2") + "");

            //workbook.Settings.CalcMode = (Aspose.Cells.CalcModeType.Manual);
            //workbook.CalculateFormula(true);
            #region 第一行

            sheet.Cells["A1"].PutValue("BRB附加有效阻尼比计算实用程序数据及计算过程Excel表 - 广东省建筑设计研究院ADG机场所");
            sheet.Cells["A2"].PutValue("计算结果");
            sheet.Cells["B2"].PutValue("楼层信息");
            sheet.Cells["C2"].PutValue("maxFi层间剪力");
            sheet.Cells["D2"].PutValue("minFi层间剪力");
            sheet.Cells["E2"].PutValue("maxUi层位移");
            sheet.Cells["F2"].PutValue("minUi层位移");
            sheet.Cells["G2"].PutValue("maxFiG层剪力");
            sheet.Cells["H2"].PutValue("-");
            sheet.Cells["I2"].PutValue("maxWs应变能");
            sheet.Cells["J2"].PutValue("minWs应变能");

            sheet.Cells["K2"].PutValue("消能器ID");
            sheet.Cells["L2"].PutValue("单圈耗能~" + Global.maxtime.ToString("F2"));
            sheet.Cells["M2"].PutValue("等效截面积");


            var s = sheet.Cells["A2"].GetStyle();
            s.ForegroundColor = System.Drawing.Color.LightGray;
            s.Pattern = BackgroundType.Solid;
            //s.Borders.DiagonalStyle = CellBorderType.Medium;
            
            //sheet.Cells["A2"].SetStyle(s);
            Range range = sheet.Cells.CreateRange("A2", "M2");
            //Cell cell = range[0, 0];
            //cell.Style.Font = 9;
            var sf = new StyleFlag();
            sf.All = true;
            range.ApplyStyle(s, sf);

            #endregion
            
            #region 循环写入内容
            int count = -4;
            int xnqCount = -4;
            int xnqTotal = 1;
            int loucengCount = 0;

            int maxCountOfFloorTotal = 0;
            int maxCountOfXnqTotal = 0;
            int zxxR1Count = 2;
            int zxxR2Count = 2;
            foreach (KeyValuePair<string,dbFloorStatus> item in db.status["CHiCHi"].floors)
            {
                if (item.Value.ifFloor == true)
                {
                    sheet = workbook.Worksheets["楼层"];
                    sheet.FreezePanes(12, 1, 12, 1);//冻结第一行第一列
                    initSheet(sheet);
                    count += 5;
                    
                    if (item.Key.IndexOf("Total") > -1) {count = loucengCount + 5;loucengCount+=5;}
                    else loucengCount = count;
                }else 
                {
                    if (xnqCount > 100)
                    {
                        xnqCount = -4;
                        xnqTotal += 1;
                    }
                    xnqCount += 5;
                    
                    count = xnqCount;
                    sheet = getSheetByName(workbook, "消能器" + xnqTotal.ToString());
                    sheet.FreezePanes(12, 1, 12, 1);//冻结第一行第一列
                    initSheet(sheet);
                    MainForm.SafeInvoke(form.listBox2, new MainForm.TaskDelegate(delegate()
                    {
                        int cur = form.listBox2.Items.Add("正在导出:"
                            + xnqTotal+"("+xnqCount+"/100)" + "/" + (form.listBox1.Items.Count / 20 -1));
                        form.listBox2.SelectedIndex = cur;
                    }));     
                }

                sheet.Cells[1, count].PutValue(item.Key);
                sheet.Cells.Merge(1, count, 1, 5);//合并单元格 

                sheet.Cells[2, count + 0].PutValue("弹性阶段k");
                sheet.Cells[2, count + 1].PutValue(item.Value.k.ToString("F1") + " " + item.Value.k1.ToString("F1"));
                sheet.Cells[3, count + 0].PutValue("maxFi");
                sheet.Cells[3, count + 1].PutValue(item.Value.maxFi);
                sheet.Cells[4, count + 0].PutValue("maxFiT");
                sheet.Cells[4, count + 1].PutValue(item.Value.maxFiT);

                sheet.Cells[5, count + 0].PutValue("maxUi");
                sheet.Cells[5, count + 1].PutValue(item.Value.maxUi);
                sheet.Cells[6, count + 0].PutValue("maxUiT");
                sheet.Cells[6, count + 1].PutValue(item.Value.maxUiT);
                sheet.Cells[7, count + 0].PutValue("滞回曲线单圈耗能");
                sheet.Cells[7, count + 1].PutValue(item.Value.zhqxArea);
                
                if (!item.Value.ifFloor) //总信息也保存单圈耗能信息
                {
                    //workbook.Worksheets["总信息"].Cells[zxxR1Count, 6].PutValue(item.Key);
                    workbook.Worksheets["总信息"].Cells[zxxR1Count, 10].Formula = "=HYPERLINK(\"#"+sheet.Name+"!"+l2n(1,count+0)+"\","+item.Key+")";
                    workbook.Worksheets["总信息"].Cells[zxxR1Count, 11].PutValue(item.Value.zhqxArea);
                    workbook.Worksheets["总信息"].Cells[zxxR1Count, 12].PutValue(meterial2Area[elementId2meterial[item.Key]].ToString());
                    zxxR1Count++;
                }
                else
                {
                    workbook.Worksheets["总信息"].Cells[zxxR2Count, 1].PutValue(item.Key);
                    workbook.Worksheets["总信息"].Cells[zxxR2Count, 2].PutValue(item.Value.maxFi);
                    workbook.Worksheets["总信息"].Cells[zxxR2Count, 3].PutValue(item.Value.minFi);
                    workbook.Worksheets["总信息"].Cells[zxxR2Count, 4].PutValue(item.Value.maxUi);
                    workbook.Worksheets["总信息"].Cells[zxxR2Count, 5].PutValue(item.Value.minUi);
                    workbook.Worksheets["总信息"].Cells[zxxR2Count, 6].PutValue(item.Value.maxFiG);
                    //workbook.Worksheets["总信息"].Cells[zxxR2Count, 7].PutValue(item.Value.minFiG);
                    workbook.Worksheets["总信息"].Cells[zxxR2Count, 8].PutValue(item.Value.maxWs);
                    workbook.Worksheets["总信息"].Cells[zxxR2Count, 9].PutValue(item.Value.minWs);
                    zxxR2Count++;
                }

                sheet.Cells[8, count + 0].PutValue("最大弹性变形");
                sheet.Cells[8, count + 1].PutValue(item.Value.tanxingD);
                sheet.Cells[9, count + 0].PutValue("全楼maxWs时刻内滞回曲线单圈耗能（四点拟合x,y）");
                sheet.Cells.Merge(9, count, 2, 1);//合并单元格 

                sheet.Cells[9, count + 1].PutValue(item.Value.d0.fx);
                sheet.Cells[9, count + 2].PutValue(item.Value.d0.ux);
                sheet.Cells[9, count + 3].PutValue(item.Value.d1.fx);
                sheet.Cells[9, count + 4].PutValue(item.Value.d1.ux);
                sheet.Cells[10, count + 1].PutValue(item.Value.d2.fx);
                sheet.Cells[10, count + 2].PutValue(item.Value.d2.ux);
                sheet.Cells[10, count + 3].PutValue(item.Value.d3.fx);
                sheet.Cells[10, count + 4].PutValue(item.Value.d3.ux);

                sheet.Cells[9, count + 1].PutValue(item.Value.zhqxAreaFxString);
                sheet.Cells[9, count + 2].PutValue(item.Value.zhqxAreaUxString);



                
                //sheet.Cells[3, count + 2].PutValue("最大弹性变形 = " + item.Value.);

                int innerCount = 11;

                int chartIndex = sheet.Charts.Add(Aspose.Cells.Charts.ChartType.ScatterConnectedByLinesWithoutDataMarker , 2, count + 2, 9, count + 5);
                Aspose.Cells.Charts.Chart chart = sheet.Charts[chartIndex];
                chart.ShowLegend = false;


                if (!item.Value.ifFloor)
                {
                    chart.PlotArea.IsAutomaticSize = false;
                    chart.PlotArea.Height = 500;
                    chart.PlotArea.Width = 5000;

                    int serID = chart.NSeries.Add(l2n(innerCount + 1, count) + ":" + l2n(chartMaxRow, count), true);
                    chart.NSeries[serID].XValues = l2n(innerCount + 1, count + 2) + ":" + l2n(chartMaxRow, count + 2);

                    //int ser2ID = chart.NSeries.Add(l2n(9, count + 1) + "," + l2n(9, count + 3) + "," + l2n(10, count + 1) + "," + l2n(10, count + 3) + "," + l2n(9, count + 1), true);
                    //chart.NSeries[ser2ID].XValues = (l2n(9, count + 2) + "," + l2n(9, count + 4) + "," + l2n(10, count + 2) + "," + l2n(10, count + 4) + "," + l2n(9, count + 2));
                    int ser2ID=0;
                    if (item.Value.zhqxAreaUxString != "")
                    {
                        ser2ID = chart.NSeries.Add("{" + item.Value.zhqxAreaUxString + "}", true);
                        chart.NSeries[ser2ID].XValues = "{" + item.Value.zhqxAreaFxString + "}";
                    }




                    chart.NSeries[serID].Smooth = false;
                    chart.NSeries[ser2ID].Smooth = false;
                }
                else
                {
                    int serID = chart.NSeries.Add(l2n(innerCount + 1, count + 4) + ":" + l2n(chartMaxRow, count + 4), true);
                }


                sheet.Cells[innerCount, count + 0].PutValue("fx");
                sheet.Cells[innerCount, count + 1].PutValue("fy");
                sheet.Cells[innerCount, count + 2].PutValue("ux");
                sheet.Cells[innerCount, count + 3].PutValue("fxG");
                sheet.Cells[innerCount, count + 4].PutValue("ws");
                foreach (KeyValuePair<double, dbPointStatus> se in item.Value.timeStat)
                {
                    innerCount ++;
                    sheet.Cells[innerCount, 0].PutValue(se.Key);
                    sheet.Cells[innerCount, count].PutValue(se.Value.fx);
                    sheet.Cells[innerCount, count + 1].PutValue(se.Value.fy);
                    sheet.Cells[innerCount, count + 2].PutValue(se.Value.ux);
                    sheet.Cells[innerCount, count + 3].PutValue(se.Value.fxG);
                    sheet.Cells[innerCount, count + 4].PutValue(se.Value.ws);
                    //string fomula = "=" + l2n(innerCount, count) + "*" + l2n(innerCount, count + 2);
                    //sheet.Cells[innerCount, count + 4].Formula = ("=" + l2n(innerCount, count) + "*" + l2n(innerCount, count + 2));
                    
                    
                    if (item.Key.IndexOf("Total") > -1){
                        sheet.Cells[innerCount, count + 4].PutValue(se.Value.ws);
                        if(item.Key=="楼层Total")maxCountOfFloorTotal=count;
                        else if(item.Key=="消能器Total") maxCountOfXnqTotal = count;
                    }


                    //sheet.Cells[innerCount, count + 4]
                    //break;
                }
                workbook.CalculateFormula(true);
                //count = count + 5;
            }

            sheet = workbook.Worksheets["楼层"];
            int znbRow = Math.Max(maxCountOfXnqTotal,maxCountOfFloorTotal)+5;
            //int znbRow = ( maxCountOfFloorTotal) + 5;
            sheet.Cells[11, znbRow].Value = "阻尼比(调整前)";
            sheet.Cells[11, znbRow+1].Value = "阻尼比(25%调整后)";
            sheet.Cells[11, znbRow + 2].Value = "取计算分段";
            sheet.Cells[10, znbRow + 1].Value = st;//开始
            sheet.Cells[10, znbRow + 2].Value = ed;// Convert.ToDouble(Global.MidasModelCalTime);//结束
            for (int i = 12; i <= Convert.ToInt32(Global.MidasModelCalTime / Global.sec)+11 ; i++)
            {
                sheet.Cells[i, znbRow].Formula = "=Abs("+ l2n(i, maxCountOfXnqTotal+4) +"/(4*3.14159*"+ 
                    l2n(i, maxCountOfFloorTotal + 4)+")*100)";
                sheet.Cells[i, znbRow + 1].Formula = "=IF(" + l2n(i, znbRow) + ">25,25,IF(" + l2n(i, znbRow) + "<=0,0," + l2n(i, znbRow) + "))";
                sheet.Cells[i, znbRow + 2].Formula = "=IF(and(" + l2n(i, 0) + ">" + l2n(10, znbRow + 1) + "," + l2n(i, 0) + "<" + l2n(10, znbRow + 2) + ")," + l2n(i, znbRow + 1) + ",\"\")";
            }
                
            int chartZnbIndex = sheet.Charts.Add(Aspose.Cells.Charts.ChartType.ScatterConnectedByLinesWithoutDataMarker, 1, znbRow, 9, znbRow+10);
            Aspose.Cells.Charts.Chart chartZnb = sheet.Charts[chartZnbIndex];
            chartZnb.ShowLegend = false;
            int nser = chartZnb.NSeries.Add(l2n(12, znbRow + 2) + ":" + l2n(Convert.ToInt32(Global.MidasModelCalTime / Global.sec) + 12, znbRow + 2), true);
            chartZnb.NSeries[nser].XValues = (l2n(12,0) + ":" + l2n(Convert.ToInt32(Global.MidasModelCalTime / Global.sec) + 12, 0));
            chartZnb.PlotArea.IsAutomaticSize = false;
            chartZnb.PlotArea.Height = 500;
            chartZnb.PlotArea.Width = 5000;
            //chartZnb.Title.Text = "附加阻尼比（时变曲线）";
            sheet.Cells[11, znbRow + 3].Formula = "=Average(" + l2n(12, znbRow + 2) + ":" + l2n(Convert.ToInt32(Global.MidasModelCalTime / Global.sec) + 12, znbRow + 2) + ")";
            workbook.Worksheets["总信息"].Cells["A4"].Formula = "=楼层!" + l2n(11, znbRow + 3);

            #endregion
            //保存
            //MainForm.lab
            workbook.Save(path+@"\"+form.listBox3.Text+"结果.xls");
            System.Diagnostics.Process.Start(path + @"\" + form.listBox3.Text + "结果.xls");

        }
    }



    public class ExcelHelper
    {
        private Excel._Application excelApp;
        private string fileName = string.Empty;
        private Excel.WorkbookClass wbclass;
        public ExcelHelper(string _filename)
        {
            excelApp = new Excel.Application();
            object objOpt = System.Reflection.Missing.Value;
            wbclass = (Excel.WorkbookClass)excelApp.Workbooks.Open(_filename, objOpt, false, objOpt, objOpt, objOpt, true, objOpt, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt);
        }
        /// <summary>
        /// 所有sheet的名称列表
        /// </summary>
        /// <returns></returns>
        public List<string> GetSheetNames()
        {
            List<string> list = new List<string>();
            Excel.Sheets sheets = wbclass.Worksheets;
            string sheetNams = string.Empty;
            foreach (Excel.Worksheet sheet in sheets)
            {
                list.Add(sheet.Name);
            }
            return list;
        }
        public Excel.Worksheet GetWorksheetByName(string name)
        {
            Excel.Worksheet sheet = null;
            Excel.Sheets sheets = wbclass.Worksheets;
            foreach (Excel.Worksheet s in sheets)
            {
                if (s.Name == name)
                {
                    sheet = s;
                    break;
                }
            }
            return sheet;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sheetName">sheet名称</param>
        /// <returns></returns>
        public Array GetContent(string sheetName)
        {
            Excel.Worksheet sheet = GetWorksheetByName(sheetName);
            //获取A1 到AM24范围的单元格
            Excel.Range rang = sheet.get_Range("K2", "L501");
            //读一个单元格内容
            //sheet.get_Range("A1", Type.Missing);
            //不为空的区域,列,行数目
            //   int l = sheet.UsedRange.Columns.Count;
            // int w = sheet.UsedRange.Rows.Count;
            //  object[,] dell = sheet.UsedRange.get_Value(Missing.Value) as object[,];
            System.Array values = (Array)rang.Cells.Value2;
            return values;
        }

        public void Close()
        {
            excelApp.Quit();
            excelApp = null;
        }

    }
}
