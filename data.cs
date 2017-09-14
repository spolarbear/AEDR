using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace sunproject
{
    public class Global
    {
        public static double sec = -1;  //间隔时间
        public static bool cal = false;
        public static double resultZnbSb, resultZnbZh, resultZnbBl;
        public static double maxtime = 0;
        public static double MidasModelCalTime = 20;  //模型时长
        public static double currentSetTime = 0;
        //public static double splitTime = -1;
    }
    public class KSet
    {
        public List<string> db;
        public KSet()
        {
            db = new List<string>();
        }
        public void add(string k)
        {
            db.Add(k);
        }
        public string get(int order)
        {
            Dictionary<string, int> amount = new Dictionary<string, int>();
            foreach (string str in db)
            {
                if (amount.ContainsKey(str)) amount[str]=amount[str]+1;
                else amount.Add(str, 1);
            }
            amount = amount.OrderByDescending(p => p.Value).ToDictionary(o => o.Key, p => p.Value);
            int i = 1;
            foreach (KeyValuePair<string,int> o in amount)
            {
                //if (Convert.ToDouble(o.Key) <0.5) continue;
                if (i == order) return o.Key;
                i++;
            }
            return "";
            //int count = db.Count(s => s == "北京");
        }
    }
    public class dbMathPoint
    {
        public static double distant(dbMathPoint p1,dbMathPoint p2){
            return Math.Sqrt((p1.x - p2.x) * (p1.x - p2.x) + (p1.y - p2.y) * (p1.y - p2.y));
        }
        public static double sBy3P(dbMathPoint p1,dbMathPoint p2,dbMathPoint p3){
            double a = dbMathPoint.distant(p1, p2);
            double b = dbMathPoint.distant(p2, p3);
            double c = dbMathPoint.distant(p3, p1);
            double p = (a + b + c) / 2;
            return Math.Sqrt( p*(p-a)*(p-b)*(p-c) )*2;
        }
        public static double sBy3P(dbPointStatus p1, dbPointStatus p2, dbPointStatus p3)
        {
            dbMathPoint mp1 = new dbMathPoint(p1.fx, p1.ux);
            dbMathPoint mp2 = new dbMathPoint(p2.fx, p2.ux);
            dbMathPoint mp3 = new dbMathPoint(p3.fx, p3.ux);
            double a = dbMathPoint.distant(mp1, mp2);
            double b = dbMathPoint.distant(mp2, mp3);
            double c = dbMathPoint.distant(mp3, mp1);
            double p = (a + b + c) / 2;
            return Math.Sqrt(p * (p - a) * (p - b) * (p - c)) * 2;
        }

        public static double zhqxEdge(Dictionary<double, dbPointStatus> pDb,double maxSec)
        {
            //return maxSec;
            double averageK = Math.Abs((pDb[maxSec].ux - pDb[Math.Round(maxSec - Global.sec, 2)].ux) / (pDb[maxSec].fx - pDb[Math.Round(maxSec - Global.sec, 2)].fx));

            for (double i = maxSec; i > 0; i = Math.Round(i - Global.sec, 2))
            {
                double currK = 0;
                double lastKey = Math.Round(i - Global.sec, 2);
                if (pDb.ContainsKey(lastKey))
                    currK = Math.Abs((pDb[i].ux - pDb[lastKey].ux) / (pDb[i].fx - pDb[lastKey].fx));
                if (Math.Abs(currK - averageK) < 0.001) 
                    averageK = (averageK + currK) / 2;
                else 
                    return i;
                //dbPointStatus gksTime = pDb[maxSec];
            }
            return 0;
        }

        public double x,y,d;
        public dbMathPoint(double x, double y)
        {
            this.x = x;
            this.y = y;
            this.d = Math.Sqrt(x * x + y * y / 10000);
        }
        public void reset(double x, double y){
            //滞回曲线外包面积 筛选：当与原点距离大于原距离，替换。
            if (Math.Sqrt(x * x + y * y / 10000) > this.d)
            {
                this.x = x;
                this.y = y;
                this.d = Math.Sqrt(x * x + y * y / 10000);
            }
        }
    }
    public class dbPointStatus//点数据
    {
        public double ux,uy,fx,fy,h,z,ws,fxG,uxG,fyG,uyG;//位移，剪力，层高，标高
        public dbPointStatus() { ux = uy = fx = fy = h = z =ws =0; }
        public dbPointStatus(double u, double f,string dir)
        {
            if (dir == "x")
            {
                this.ux = u;
                this.fx = f;
            }
            else
            {
                this.uy = u;
                this.fy = f;
            }
        }
        public double distTo(dbPointStatus anotherPoint)
        {
            return Math.Sqrt((ux - anotherPoint.ux) * (ux - anotherPoint.ux) + (uy - anotherPoint.uy) * (uy - anotherPoint.uy));
        }
    }
    public class dbFloorStatus//楼层或单个消能器数据
    {
        public dbPointStatus floorPoint = new dbPointStatus();//本层数据
        public Dictionary<double, dbPointStatus> timeStat=new Dictionary<double,dbPointStatus> ();//时程点集
        public void pAdd(double name, dbPointStatus dbp)
        {
            if (!timeStat.ContainsKey(name))
                timeStat.Add(name, dbp);
        }


        public int totalFX = 0, totalFY = 0, totalUX = 0, totalUY = 0;
        public double tanxingD = 0;
        public dbPointStatus tanxinD1 = new dbPointStatus(), tanxinD2 = new dbPointStatus();
        public dbPointStatus d0 = new dbPointStatus(), d1 = new dbPointStatus(), d2 = new dbPointStatus(), d3 = new dbPointStatus();
        public double zhqxArea = 0;
        public string zhqxAreaUxString = "";
        public string zhqxAreaFxString = "";
        public double k = 0,k1=0; public double maxy = 0; public double miny = 0;
        public bool ifFloor=true;
        public double maxFi = 0, maxFiT = 0, minFi = 0, minFiT = 0, maxUi = 0, maxUiT = 0, minUi = 0, minUiT = 0, maxWs = 0, maxWsT = 0, maxFiG = 0, maxFiGT = 0,minWs=0,minWsT=0;//包络计算方法
        public double maxFx = 0, minFx = 0;
        public double maxFiG_CutTopFloor_maxFiG = 0;
        public dbFloorStatus()
        {    }
        public dbFloorStatus(dbPointStatus ps)
        {
            floorPoint = ps;
        }
    }
    public class dbStatus//工况模型
    {
        public Dictionary<string,dbFloorStatus> floors=new Dictionary<string,dbFloorStatus> ();//各楼层
        public dbStatus()
        {     }
        public dbStatus(string name)
        {
            if(!floors.ContainsKey(name))
                floors.Add(name, new dbFloorStatus());//初始化楼层
        }
    }
    public class db//全楼模型
    {
        public Dictionary<string,dbStatus> status = new Dictionary<string,dbStatus> ();//各工况集
        //eg.   [CHiChi]  [人工波]
        public double timePreSet = 0;
       
    }
}
