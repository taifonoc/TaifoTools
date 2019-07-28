using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 多重頁面
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            UrlAddress N = new UrlAddress();
            N.Urladdress("https://branch.taipower.com.tw/Content/NoticeBlackout/bulletin.aspx?&SiteID=564732646551216421&MmmID=616371300113254267");
            N.area("北市區");
            N.Html_Original();



            UrlAddress NN = new UrlAddress();
            NN.Urladdress("https://branch.taipower.com.tw/Content/NoticeBlackout/bulletin.aspx?&SiteID=564732636524040174&MmmID=616371300130136031");
            NN.area("北北區");
            NN.Html_Original();


            UrlAddress NS = new UrlAddress();
            NS.Urladdress("https://branch.taipower.com.tw/Content/NoticeBlackout/bulletin.aspx?&SiteID=564732646356736245&MmmID=616371300115522273");
            NS.area("北南區");
            NS.Html_Original();


            UrlAddress NW = new UrlAddress();
            NW.Urladdress("https://branch.taipower.com.tw/Content/NoticeBlackout/bulletin.aspx?&SiteID=564766277367364243&MmmID=616371300000777256");
            NW.area("北西區");
            NW.Html_Original();



            Application.Run(new Form1());
        }
    }
}
