using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Net;
using CsQuery;

using NPOI;
using NPOI.HSSF;
using NPOI.HSSF.Util;
using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Data;
using System.Data.SqlClient;
using HtmlAgilityPack;
using System.Threading.Tasks;
using System.Linq;

namespace Brian_Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Test By Brian.");
            Console.WriteLine("(1)爬蟲取得NBA隊名");
            Console.WriteLine("(2)爬蟲");
            Console.WriteLine("(3)Excel");
            Console.WriteLine("(4)");
            Console.WriteLine("(5)");
            Console.WriteLine("(6)");
            Console.WriteLine("(7)");
            Console.WriteLine("(8)");
            Console.WriteLine("(9)");
            Console.WriteLine("(10)");
            Console.Write("輸入想執行的方法數字: ");
            string input = Console.ReadLine();

            switch (input)
            {
                case "1":
                    getNBATeamName();
                    break;
                case "2":
                    getDetail();
                    break;
                case "3":
                    //testOutExcel();
                    break;
                case "4":
                    break;
                case "5":
                    break;
                case "6":
                    break;
                case "7":
                    break;
                case "8":
                    break;
                case "9":
                    break;
                case "10":
                    break;
            }


            Console.WriteLine("Finish.");

            string str = Console.ReadLine();
        }


        static void getDetail()
        {
            string htmlContent = GetContent("https://www.cwb.gov.tw/V7/index.htm#self");
            List<string> ls = new List<string>();

            foreach (string item in GetHtmlBySelector("table.BoxTable tbody tr td a", htmlContent))
            {
                Console.WriteLine(item);

                //int indexOfStart = item.IndexOf(getAttribute) + attributeLength; //篩到屬性名稱
                //indexOfStart = indexOfStart + item.Substring(indexOfStart).IndexOf("\"") + 1; // 篩到第一個 "
                //string str = item.Substring(indexOfStart);
                //int indexOfEnd = str.IndexOf("\""); //篩到第二個 "
                //str = item.Substring(indexOfStart, indexOfEnd); //取得屬性的值

                //ls.Add(str);
            }
        }

        //取得隊名
        static void getNBATeamName()
        {
            List<string> result = WebCrawler("https://tw.global.nba.com/standings/", "#menu_body ul li a", "value");
            Boolean isTeam = false;
            int i = 0;
            foreach (string item in result)
            {
                if (item.Contains("塞爾蒂克")) isTeam = true;
                if (!isTeam) continue;
                Console.WriteLine(item);
                i++;
                if (item.Contains("國王")) break;
            }
            Console.WriteLine(i);
        }

        //爬蟲，利用jQuery去抓url上面的內容
        static List<string> WebCrawler(string url, string jQuerySelect, string getAttribute)
        {
            string htmlContent = GetContent(url);
            int attributeLength = getAttribute.Length;
            List<string> ls = new List<string>();

            if (getAttribute.ToLower().Equals("value"))
            {
                foreach (string item in GetHtmlBySelector(jQuerySelect, htmlContent))
                {
                    int indexOfStart = item.IndexOf(">") + 1; //篩到 >
                    string str = item.Substring(indexOfStart);
                    int indexOfEnd = str.IndexOf("<"); //篩到 <
                    str = item.Substring(indexOfStart, indexOfEnd); //取得屬性的值

                    ls.Add(str);
                }
            }
            else
            {
                foreach (string item in GetHtmlBySelector(jQuerySelect, htmlContent))
                {
                    int indexOfStart = item.IndexOf(getAttribute) + attributeLength; //篩到屬性名稱
                    indexOfStart = indexOfStart + item.Substring(indexOfStart).IndexOf("\"") + 1; // 篩到第一個 "
                    string str = item.Substring(indexOfStart);
                    int indexOfEnd = str.IndexOf("\""); //篩到第二個 "
                    str = item.Substring(indexOfStart, indexOfEnd); //取得屬性的值

                    ls.Add(str);
                }
            }



            return ls;
        }

        //jQuery 部分
        public static List<string> GetHtmlBySelector(string Selector, string Html)
        {
            List<string> targets = new List<string>();
            CQ cq = CQ.Create(Html);
            foreach (IDomObject obj in cq.Find(Selector))
            {
                targets.Add(System.Net.WebUtility.HtmlDecode(obj.Render()));
            }
            return targets;
        }

        //抓HTML內容部分
        public static string GetContent(string Url)
        {
            string Content = string.Empty, Title = string.Empty;
            try
            {
                HttpWebRequest request = WebRequest.Create(Url) as HttpWebRequest;
                request.Credentials = CredentialCache.DefaultNetworkCredentials;
                if (Url.StartsWith("https"))
                {
                    ServicePointManager.SecurityProtocol = (SecurityProtocolType)4032; ;
                    ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                }
                request.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
                request.PreAuthenticate = true;
                request.AllowAutoRedirect = true;
                request.MaximumAutomaticRedirections = 100;
                request.Timeout = 20 * 1000;
                request.Accept = "*/*";
                request.Headers.Add("Accept-Language", "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7");
                request.Headers.Add("Accept-Encoding", "gzip, deflate, br");
                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36";
                request.Method = "GET";
                request.KeepAlive = true;
                request.Host = new Uri(Url).Host;

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                string ContentType = response.ContentType.ToLower();
                string CharacterSet = response.CharacterSet;

                if (ContentType.ToLower().Contains("text"))
                {
                    byte[] Data = new byte[0];
                    Stream responseStream = response.GetResponseStream();
                    using (MemoryStream ms = new MemoryStream())
                    {
                        byte[] buffer = new byte[4096];
                        int Length = responseStream.Read(buffer, 0, buffer.Length);
                        while (Length > 0)
                        {
                            ms.Write(buffer, 0, Length);
                            Length = responseStream.Read(buffer, 0, buffer.Length);
                        }
                        Data = ms.ToArray();
                    }

                    Content = Encoding.UTF8.GetString(Data, 0, Data.Length);

                    if (Content.ToLower().Contains("charset=big5") || CharacterSet.ToLower().Contains("big5"))
                    {
                        Content = Encoding.GetEncoding("big5").GetString(Data, 0, Data.Length);
                    }
                }
                response.Close();
                request.Abort();
            }
            catch (WebException)
            {
                if (Url.StartsWith("http:"))
                {
                    return GetContent(Url.Replace("http:", "https:"));
                }
                else Content = "[[404]]";
            }
            catch (UriFormatException)
            {
                Content = "[[網址不正確]]";
            }
            catch (Exception ex)
            {
                Content = "[[" + ex.Message + "]]";
            }
            finally { }

            return Content;
        }
    }
}
