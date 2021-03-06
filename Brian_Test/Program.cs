﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Net;
using CsQuery;

using NPOI;
using NPOI.HSSF;
using NPOI.XSSF;
using NPOI.HSSF.Util;
using NPOI.XSSF.Util;
using NPOI.HSSF.Model;
using NPOI.XSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
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
            Console.WriteLine("(4)上方導覽");
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
                    getSitemap();
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


        public static void NPOIToExcel(List<List<List<string>>> ls,string siteName)
        {

            IWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet1 = (HSSFSheet)wb.CreateSheet(siteName);

            //HSSFWorkbook wb = new HSSFWorkbook();
            //HSSFSheet sheet1 = (HSSFSheet)wb.CreateSheet("權限表");

            MemoryStream ms = new MemoryStream();
            //MemoryStream ms = new MemoryStream();

            HSSFRow row = null;
            //HSSFRow row = null;

            //設定儲存格樣式
            HSSFCell cell = null;
            HSSFCellStyle wrapStyle = (HSSFCellStyle)wb.CreateCellStyle();

            ////設定儲存格樣式
            //HSSFCell cell = null;
            //HSSFCellStyle wrapStyle = null;
            ////HSSFCellStyle wrapStyle10 = null;
            ////HSSFCellStyle wrapStyleR10 = null;
            ////HSSFCellStyle colorStyle = null;
            ////HSSFCellStyle RightStyle = null;
            ////HSSFCellStyle CenterStyle = null;
            //wrapStyle = (HSSFCellStyle)wb.CreateCellStyle();
            ////wrapStyle10 = (HSSFCellStyle)wb.CreateCellStyle();
            ////wrapStyleR10 = (HSSFCellStyle)wb.CreateCellStyle();
            ////colorStyle = (HSSFCellStyle)wb.CreateCellStyle();
            ////RightStyle = (HSSFCellStyle)wb.CreateCellStyle();

            HSSFFont font1 = (HSSFFont)wb.CreateFont();
            //字體尺寸
            font1.FontHeightInPoints = 12;

            //HSSFFont font1 = (HSSFFont)wb.CreateFont();
            ////字體尺寸
            //font1.FontHeightInPoints = 12;

            //HSSFFont font10 = (HSSFFont)wb.CreateFont();
            ////字體尺寸
            //font10.FontHeightInPoints = 10;
            //font10.FontName = "細明體";


            ////RightStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            ////RightStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            ////RightStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            ////RightStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            ////RightStyle.Alignment = HorizontalAlignment.Right;
            ////RightStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("#,##0");
            //////RightStyle.SetFont(font1);

            ////CenterStyle = (HSSFCellStyle)wb.CreateCellStyle();
            ////CenterStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            ////CenterStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            ////CenterStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            ////CenterStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            ////CenterStyle.Alignment = HorizontalAlignment.Center;
            ////CenterStyle.WrapText = true;
            ////CenterStyle.VerticalAlignment = VerticalAlignment.Center;
            ////CenterStyle.SetFont(font1);

            wrapStyle.SetFont(font1);
            wrapStyle.WrapText = true;
            wrapStyle.BorderTop = BorderStyle.Thin;
            wrapStyle.BorderLeft = BorderStyle.Thin;
            wrapStyle.BorderBottom = BorderStyle.Thin;
            wrapStyle.BorderRight = BorderStyle.Thin;
            wrapStyle.FillForegroundColor = HSSFColor.Red.Index;
            wrapStyle.VerticalAlignment = VerticalAlignment.Center;

            ////wrapStyle.SetFont(font1);
            //wrapStyle.WrapText = true;
            //wrapStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            //wrapStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            //wrapStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            //wrapStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            //wrapStyle.FillForegroundColor = HSSFColor.Red.Index;
            //wrapStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;

            ////wrapStyle10.SetFont(font10);
            ////wrapStyle10.WrapText = true;
            ////wrapStyle10.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            ////wrapStyle10.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            ////wrapStyle10.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            ////wrapStyle10.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            ////wrapStyle10.FillForegroundColor = HSSFColor.Red.Index;
            ////wrapStyle10.VerticalAlignment = VerticalAlignment.Center;


            ////wrapStyleR10.SetFont(font10);
            ////wrapStyleR10.WrapText = true;
            ////wrapStyleR10.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            ////wrapStyleR10.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            ////wrapStyleR10.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            ////wrapStyleR10.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            ////wrapStyleR10.FillForegroundColor = HSSFColor.Red.Index;
            ////wrapStyleR10.Alignment = HorizontalAlignment.Right;
            ////wrapStyleR10.VerticalAlignment = VerticalAlignment.Center;

            ////colorStyle.FillPattern = FillPattern.SolidForeground;
            ////colorStyle.FillBackgroundColor = HSSFColor.Red.Index;
            ////colorStyle.FillForegroundColor = 10;
            sheet1.PrintSetup.Landscape = true;
            sheet1.ForceFormulaRecalculation = true;

            int rowIndex = 0;
            int field = 0;
            int rowPartIndex = 0;

            row = (HSSFRow)sheet1.CreateRow(rowIndex);
            cell = (HSSFCell)row.CreateCell(0);
            cell.CellStyle = wrapStyle;  //指定樣式
            cell.SetCellType(CellType.String);
            cell.SetCellValue("頁面名稱");
            cell = (HSSFCell)row.CreateCell(1);
            cell.CellStyle = wrapStyle;  //指定樣式
            cell.SetCellType(CellType.String);
            cell.SetCellValue("連結");
            rowIndex++;
            rowPartIndex++;

            foreach (List<List<string>> part in ls)
            {
                int size = part.First().Count;
                for (int i = rowPartIndex; i < size; i++)
                {
                    row = (HSSFRow)sheet1.CreateRow(i);
                }

                field = 0;

                foreach (List<string> part1 in part)
                {
                    rowIndex = rowPartIndex;
                    foreach (string str in part1)
                    {
                        row = (HSSFRow)sheet1.GetRow(rowIndex);
                        cell = (HSSFCell)row.CreateCell(field);
                        cell.CellStyle = wrapStyle;  //指定樣式
                        //cell.SetCellType(CellType.String);
                        cell.SetCellValue(str);
                        rowIndex++;
                    }
                    field++;
                }
                rowPartIndex = rowIndex;
            }

            //DataTable dt = pDataTable; //準備好要寫入的資料
            //int rowIndex = 0;
            //int dtSize = dt.Rows.Count;


            //row = (HSSFRow)sheet1.CreateRow(rowIndex);
            //cell = (HSSFCell)row.CreateCell(0);
            //cell.CellStyle = wrapStyle;  //指定樣式
            //cell.SetCellType(CellType.String);
            //cell.SetCellValue("站台");
            //cell = (HSSFCell)row.CreateCell(1);
            //cell.CellStyle = wrapStyle;  //指定樣式
            //cell.SetCellType(CellType.String);
            //cell.SetCellValue("單位");
            //cell = (HSSFCell)row.CreateCell(2);
            //cell.CellStyle = wrapStyle;  //指定樣式
            //cell.SetCellType(CellType.String);
            //cell.SetCellValue("帳號");
            //cell = (HSSFCell)row.CreateCell(3);
            //cell.CellStyle = wrapStyle;  //指定樣式
            //cell.SetCellType(CellType.String);
            //cell.SetCellValue("姓名");
            //cell = (HSSFCell)row.CreateCell(4);
            //cell.CellStyle = wrapStyle;  //指定樣式
            //cell.SetCellType(CellType.String);
            //cell.SetCellValue("系統權限");
            //rowIndex++;

            //try
            //{
            //    for (int i = 0; i < dtSize; i++) //逐筆資料寫入
            //    {

            //        string field1 = dt.Rows[i]["站台"].ToString().Trim();
            //        string field2 = dt.Rows[i]["單位"].ToString().Trim();
            //        string field3 = dt.Rows[i]["帳號"].ToString().Trim();
            //        string field4 = dt.Rows[i]["姓名"].ToString().Trim();
            //        string field5 = dt.Rows[i]["權限"].ToString().Trim();

            //        row = (HSSFRow)sheet1.CreateRow(rowIndex); //新的一行
            //        cell = (HSSFCell)row.CreateCell(0); //第幾列 站台
            //        cell.CellStyle = wrapStyle;  //指定樣式
            //        cell.SetCellType(CellType.String); //設定欄位格式
            //        cell.SetCellValue(field1);   //寫入值
            //        cell = (HSSFCell)row.CreateCell(1); //單位
            //        cell.CellStyle = wrapStyle;
            //        cell.SetCellType(CellType.String);
            //        cell.SetCellValue(field2);
            //        cell = (HSSFCell)row.CreateCell(2); //帳號
            //        cell.CellStyle = wrapStyle;
            //        cell.SetCellType(CellType.String);
            //        cell.SetCellValue(field3);
            //        cell = (HSSFCell)row.CreateCell(3); //姓名
            //        cell.CellStyle = wrapStyle;
            //        cell.SetCellType(CellType.String);
            //        cell.SetCellValue(field4);
            //        cell = (HSSFCell)row.CreateCell(4); //系統權限
            //        cell.CellStyle = wrapStyle;
            //        if (field5.Length > 30000)
            //        {
            //            cell.SetCellType(CellType.String);
            //            cell.SetCellValue("系統管理者");
            //        }
            //        else
            //        {
            //            cell.SetCellType(CellType.String);
            //            cell.SetCellValue(field5);
            //        }

            //        rowIndex++;
            //    }

            //}
            //catch (Exception e)
            //{
            //    sheet1.GetRow(1).GetCell(0).SetCellValue("發生錯誤：" + e.Message);
            //}
            sheet1.SetColumnWidth(0, 20 * 256);
            sheet1.SetColumnWidth(1, 30 * 256);
            //sheet1.SetColumnWidth(2, 20 * 256);
            //sheet1.SetColumnWidth(4, 30 * 256);

            //產生檔案
            //FileStream FS = new FileStream(Path.Combine("D:\\Brian\\", siteName+".xlsx"), FileMode.Create, System.IO.FileAccess.Write);
            FileStream FS = File.Create(siteName+".xlsx");
            wb.Write(FS);
            FS.Close();

            ////產生下載的檔案
            //wb.Write(ms);
            //wb = null;
            //ms.Close();
            //ms.Dispose();
            //oPage.Response.Buffer = true;
            //oPage.Response.Clear();
            //oPage.Response.ContentType = "application/octet-stream";
            //oPage.Response.AddHeader("Content-Disposition", "attachment;filename=\"" + FileName + ".xls\"");
            //oPage.Response.BinaryWrite(ms.ToArray());
            //oPage.Response.Flush();
            //oPage.Response.End();
        
        }

        static void getSitemap()
        {
            Console.Write("請輸入該網站的網站導覽網址:");
            string url = Console.ReadLine();
            string htmlContent = GetContent(url);
            List<List<List<string>>> ls = new List<List<List<string>>>(); 
            //foreach (string item in GetHtmlBySelector("div form a", htmlContent))
            //{
            //    Console.WriteLine(item);
            //}
            List<string> topText = WebCrawler(url, "div form a", "value");
            List<string> topLink = WebCrawler(url, "div form a", "href");
            foreach (string item in topText)
            {
                Console.WriteLine(item);
            }
            foreach (string item in topLink)
            {
                Console.WriteLine(item);
            }
            List<List<string>> top = new List<List<string>>();
            top.Add(topText);
            top.Add(topLink);
            //foreach (string item in GetHtmlBySelector("ul li a", htmlContent))
            //{
            //    Console.WriteLine(item);
            //}
            List<string> contentText = WebCrawler(url, "ul li a", "value");
            List<string> contentLink = WebCrawler(url, "ul li a", "href");
            foreach (string item in contentText)
            {
                Console.WriteLine(item);
            }
            foreach (string item in contentLink)
            {
                Console.WriteLine(item);
            }
            List<List<string>> content = new List<List<string>>();
            content.Add(contentText);
            content.Add(contentLink);
            
            //foreach (string item in GetHtmlBySelector("a", htmlContent))
            //{
            //    foreach (string str in GetHtmlBySelector("a img", htmlContent))
            //    {
            //        if (item.Contains(str))
            //        {
            //            Console.WriteLine(getAttribute(item, "href"));
            //            break;
            //        }
            //    }
            //}
            //foreach (string item in GetHtmlBySelector("a img", htmlContent))
            //{
            //    Console.WriteLine(getAttribute(item, "alt"));
            //}
            //getAttribute();
            List<string> footAlt = new List<string>();
            List<string> footLink = new List<string>();
            foreach (string item in GetHtmlBySelector("a", htmlContent))
            {
                foreach (string str in GetHtmlBySelector("a img", htmlContent))
                {
                    if (item.Contains(str))
                    {
                        Console.WriteLine(getAttribute(item, "href"));
                        footLink.Add(getAttribute(item, "href"));
                        break;
                    }
                }
            }
            foreach (string item in GetHtmlBySelector("a img", htmlContent))
            {
                Console.WriteLine(getAttribute(item, "alt"));
                footAlt.Add(getAttribute(item, "alt"));
            }
            List<List<string>> foot = new List<List<string>>();
            foot.Add(footAlt);
            foot.Add(footLink);

            ls.Add(top);
            ls.Add(content);
            ls.Add(foot);

            string siteName = getAttribute(GetHtmlBySelector("title", htmlContent).Where(o => o == o).First(), "value");

            //foreach (string item in GetHtmlBySelector("title", htmlContent))
            //{
            //    Console.WriteLine(item);
            //}

            NPOIToExcel(ls, siteName);
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

        public static string getAttribute(string result,string attr)
        {
            int attributeLength = attr.Length;
            string str = string.Empty;
            if (attr.ToLower().Equals("value"))
            {
                int indexOfStart = result.IndexOf(">") + 1; //篩到 >
                str = result.Substring(indexOfStart);
                int indexOfEnd = str.IndexOf("<"); //篩到 <
                str = result.Substring(indexOfStart, indexOfEnd); //取得屬性的值
            }
            else
            {

                int indexOfStart = result.IndexOf(attr) + attributeLength; //篩到屬性名稱
                indexOfStart = indexOfStart + result.Substring(indexOfStart).IndexOf("\"") + 1; // 篩到第一個 "
                str = result.Substring(indexOfStart);
                int indexOfEnd = str.IndexOf("\""); //篩到第二個 "
                str = result.Substring(indexOfStart, indexOfEnd); //取得屬性的值
            }
            return str;
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
