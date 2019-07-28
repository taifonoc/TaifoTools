using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;


namespace 多重頁面
{
    class UrlAddress
    {
        private string Url { get; set; }
        private string Area { get; set; }
        //private bool SW { get; set; }
        //public void Sw(bool sw)
        //{
        //    SW = sw;
        //}

        public void Urladdress(string urlAddress)
        {
            Url = urlAddress;
        }
        public void area(string aRea)
        {
            Area = aRea;
        }
        public void Html_Original()
        {
            //string[] DATA = new string[5000];
            // int count = 0;

            HttpWebRequest Request = (HttpWebRequest)WebRequest.Create(Url);
            HttpWebResponse Response = (HttpWebResponse)Request.GetResponse();
            if (Response.StatusCode == HttpStatusCode.OK)
            {
                Stream ReceiveStream = Response.GetResponseStream();
                StreamReader ReadStream = null;
                if (Response.CharacterSet == null)
                {
                    ReadStream = new StreamReader(ReceiveStream);
                }
                else
                {
                    ReadStream = new StreamReader(ReceiveStream, Encoding.GetEncoding(Response.CharacterSet));
                }

                string Data = ReadStream.ReadToEnd().Replace("，", Environment.NewLine);

                HtmlDocument Doc = new HtmlDocument();
                Doc.LoadHtml(Data);



                //  if (SW == true)
                // {


                DateTime d = DateTime.Now;
                string Sameday = d.GetDateTimeFormats('D')[1].ToString();
                string dirPath = $@"{AppDomain.CurrentDomain.BaseDirectory}\{Area}";
                if (Directory.Exists(dirPath))
                {
                    Console.WriteLine("success");
                }
                else
                {
                    Directory.CreateDirectory(dirPath);
                    Console.WriteLine("The directory {0} was created.", dirPath);
                }

                StreamWriter str = new StreamWriter($@"{AppDomain.CurrentDomain.BaseDirectory}\{Area}\{Sameday}-{Area}.txt");
                // str.WriteLine(data);
                if (Doc.DocumentNode.SelectNodes("//table") != null)
                {

                    foreach (HtmlNode table in Doc.DocumentNode.SelectNodes("//table"))
                    {
                        str.WriteLine($"\n({Area})日期:" + table.Id + "\n");
                        //DATA[count] = $"\n({Area})日期:" + table.Id + "\n";
                        //count++;
                        //Console.WriteLine("Date:" + table.Id);
                        foreach (HtmlNode row in table.SelectNodes("tr"))
                        {

                            str.WriteLine();
                            //DATA[count] = $"\n";
                            //count++;
                            //Console.WriteLine("row");
                            foreach (HtmlNode cell in row.SelectNodes("th|td"))
                            {

                                str.WriteLine(cell.InnerText);
                                //DATA[count] = $"\n" + cell.InnerText + "\n";
                                //count++;
                                // Console.WriteLine("cell:" + cell.InnerText);
                            }
                        }
                    }

                    str.Close();

                    //}

                }
                ReadStream.Close();
            }

            Response.Close();
        }


    }
}
