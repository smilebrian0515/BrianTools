using System;
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

namespace TestMakeTools
{
    class Program
    {
        static void Main(string[] args)
        {
            string input = "0";
            while (!input.Equals("-1"))
            {
                Console.WriteLine("This Tool is By Brian.");
                Console.WriteLine("(1)爬蟲取得NBA隊名");
                Console.WriteLine("(2)爬蟲");
                Console.WriteLine("(3)Excel");
                Console.WriteLine("(4)網站導覽，輸入一個網址");
                Console.WriteLine("(5)網站導覽，一次輸入一整串，輸入-1後結束輸入，開始產生EXCEL");
                Console.WriteLine("(6)爬蟲(未完成)");
                Console.WriteLine("(7)串HTML語法");
                Console.WriteLine("(8)比對EXCEL內容");
                Console.WriteLine("(9)");
                Console.WriteLine("(10)");
                Console.Write("輸入想執行的方法數字(輸入-1離開): ");
                input = Console.ReadLine();
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
                        getSitemapForTainan("", "div form a", "#sitecontent ul li a", "");
                        break;
                    case "5":
                        List<string> URLls = new List<string>();
                        Console.Write("請輸入網站網址: ");
                        string url = Console.ReadLine();
                        while (!url.Equals("-1"))
                        {
                            URLls.Add(url);
                            Console.Write("請輸入網站網址: ");
                            url = Console.ReadLine();
                        }
                        foreach (string item in URLls)
                        {
                            getSitemapForTainan(item, "div form a", "#sitecontent ul li a", "");
                        }
                        break;
                    case "6":
                        Console.Write("請輸入上方導覽的Query指令: ");
                        string topQuery = Console.ReadLine();
                        Console.Write("請輸入內容的Query指令: ");
                        string contentQuery = Console.ReadLine();
                        Console.Write("請輸入下方導覽的Query指令: ");
                        string foot = Console.ReadLine();
                        break;
                    case "7":
                        make0419();
                        break;
                    case "8":
                        toVictoria();
                        break;
                    case "9":
                        break;
                    case "10":
                        break;
                }
                Console.Write("\n\n\n\n\n");
            }

            Console.WriteLine("Finish.");
            string str = Console.ReadLine();
        }

        public static int getExcelFieldIndex(string input)
        {
            string str = input;
            str = str.ToUpper();
            if (str.Length == 1)
            {
                char c = (char)str.FirstOrDefault();
                return (((int)c) - 65);
            }
            else
            {
                if (str.Length > 2)
                {
                    throw new System.Exception("很抱歉，目前不支援ZZ以後的欄位，有需要請自行修改，或通知程式撰寫者。");
                }
                char[] arr = str.ToArray();
                int i = ((((int)arr[0]) - 64) * 26) + (((int)arr[1]) - 65);
                return i;
            }
        }


        public static void toVictoria()
        {
            Console.WriteLine("請輸入兩個Excel檔名，將會為您比較兩個檔案內的資料，請務必輸入正確檔名(並包含副檔名)");
            Console.Write("請輸入第一個檔名(EX:test.xlsx):");
            string srcFile = Console.ReadLine();
            Console.Write("請輸入第二個檔名(EX:test2.xlsx):");
            string srcFile2 = Console.ReadLine();
            DataTable dt = loadExcel(srcFile);
            DataTable dt2 = loadExcel(srcFile2);

            //DataTable dt = loadExcel("test.xlsx");
            //DataTable dt2 = loadExcel("test2.xlsx");
            int cellSize = dt.Columns.Count;
            //List<int> fieldList = new List<int>();
            //使用者的需求欄位
            //int debugField1 = 1; //3
            //int debugField2 = 2; //7
            //int debugField3 = 3; //18
            List<int> needField = new List<int>();
            string inputField = string.Empty;
            Console.WriteLine("請輸入要比較的欄位，例如:D、H、AA，一次輸入一個");
            Console.WriteLine("第一個盡量為人名、ID等，比較不易變動的資料，本程式會以第一個欄位為基準去比較其他欄位的資料，來得知那些資料有變動。");
            Console.WriteLine("輸入完畢，請再按一下Enter");
            inputField = Console.ReadLine();
            while (!string.IsNullOrEmpty(inputField))
            {
                needField.Add(getExcelFieldIndex(inputField));
                inputField = Console.ReadLine();
            }

            if (needField.Count == 0)
            {
                Console.WriteLine("沒有輸入任何欄位");
                return;
            }

            int checkRange = 5;




            if (cellSize != dt2.Columns.Count)
            {
                Console.WriteLine("兩個檔案的欄位數量不一致，請確認一下EXCEL，並通知程式撰寫者。");
                return;
            }
            Boolean checkFile = true;
            for (int i = 0; i < cellSize; i++)
            {
                if (!dt.Columns[i].ColumnName.Equals(dt2.Columns[i].ColumnName))
                {
                    checkFile = false;
                    break;
                }
            }
            if (!checkFile)
            {
                Console.WriteLine("兩個檔案的欄位內容不一致，請確認一下EXCEL，並通知程式撰寫者。");
                Console.WriteLine("test.xlsx\t\ttest2.xlsx");
                for (int i = 0; i < cellSize; i++)
                {
                    Console.WriteLine(dt.Columns[i].ColumnName + "\t\t" + dt2.Columns[i].ColumnName);
                }
                return;
            }

            DataTable NewData = new DataTable();
            DataTable updateData = new DataTable();
            DataTable deleteData = new DataTable();

            for (int i = 0; i < cellSize; i++)
            {
                NewData.Columns.Add(dt.Columns[i].ColumnName, typeof(String));
                NewData.Columns[dt.Columns[i].ColumnName].MaxLength = 50;
                NewData.Columns[dt.Columns[i].ColumnName].AllowDBNull = false;

                updateData.Columns.Add(dt.Columns[i].ColumnName, typeof(String));
                updateData.Columns[dt.Columns[i].ColumnName].MaxLength = 50;
                updateData.Columns[dt.Columns[i].ColumnName].AllowDBNull = false;

                deleteData.Columns.Add(dt.Columns[i].ColumnName, typeof(String));
                deleteData.Columns[dt.Columns[i].ColumnName].MaxLength = 50;
                deleteData.Columns[dt.Columns[i].ColumnName].AllowDBNull = false;
            }

            int dt1Index = 0;
            int dt2Index = 0;
            
            while (!(dt.Rows.Count == dt1Index  || dt2.Rows.Count == dt2Index )) //還沒到底前的處理
            {
                if (checkDataEquals(dt, dt2, dt1Index, dt2Index, needField)) //資料相同
                {
                    dt1Index++;
                    dt2Index++;
                    GC.Collect();
                }
                else
                {
                    if (dt.Rows[dt1Index].ItemArray[needField.First()].ToString().Equals(dt2.Rows[dt2Index].ItemArray[needField.First()].ToString())) //資料更新
                    {
                        DataRow dtRow = updateData.NewRow();
                        dataFill(dt2, dtRow, dt2Index, 0);
                        updateData.Rows.Add(dtRow);
                        dt1Index++;
                        dt2Index++;
                        GC.Collect();
                        continue;
                    }
                    int addIndexNum = 0; //資料新增
                    for (int i = 1; i <= checkRange; i++) //檢查新增幾筆資料
                    {
                        if (dt2Index + i == dt2.Rows.Count) //到底了
                        {
                            break;
                        }
                        if (dt.Rows[dt1Index].ItemArray[needField.First()].ToString().Equals(dt2.Rows[dt2Index+i].ItemArray[needField.First()].ToString()))
                        {
                            addIndexNum = i;
                            break;
                        }
                    }
                    for (int i = 0; i < addIndexNum; i++) //將資料丟到DT
                    {
                        DataRow dtRow = NewData.NewRow();
                        dataFill(dt2, dtRow, dt2Index, i);
                        NewData.Rows.Add(dtRow);
                    }
                    if (addIndexNum > 0)
                    {
                        dt2Index += addIndexNum;
                        GC.Collect();
                        continue;
                    }

                    int delIndexNum = 0; //資料刪除
                    for (int i = 1; i <= checkRange; i++) //確認刪除幾筆資料
                    {
                        if (dt1Index + i == dt.Rows.Count) //到底了
                        {
                            break;
                        }
                        if (dt2.Rows[dt2Index].ItemArray[needField.First()].ToString().Equals(dt.Rows[dt1Index + i].ItemArray[needField.First()].ToString()))
                        {
                            delIndexNum = i;
                            break;
                        }
                    }
                    for (int i = 0; i < delIndexNum; i++) //將資料丟到DT
                    {
                        DataRow dtRow = deleteData.NewRow();
                        dataFill(dt, dtRow, dt1Index, i);
                        deleteData.Rows.Add(dtRow);
                    }
                    if (delIndexNum > 0)
                    {
                        dt1Index += delIndexNum;
                        GC.Collect();
                        continue;
                    }

                    DataRow DelRow = deleteData.NewRow();
                    dataFill(dt, DelRow, dt1Index, 0);
                    deleteData.Rows.Add(DelRow);

                    DataRow AddRow = NewData.NewRow();
                    dataFill(dt2, AddRow, dt2Index, 0);
                    NewData.Rows.Add(AddRow);
                    dt1Index++;
                    dt2Index++;
                    GC.Collect();
                }
            }
            GC.Collect();

            if (!(dt.Rows.Count == dt1Index && dt2.Rows.Count == dt2Index)) //確認是否到底
            {
                if (dt.Rows.Count == dt1Index) //新增資料
                {
                    while (!(dt2.Rows.Count == dt2Index))
                    {
                        DataRow dtRow = NewData.NewRow();
                        dataFill(dt2, dtRow, dt2Index, 0);
                        NewData.Rows.Add(dtRow);
                        dt2Index++;
                    }
                }

                if (dt2.Rows.Count == dt2Index) //刪除資料
                {
                    while (!(dt.Rows.Count == dt1Index))
                    {
                        DataRow dtRow = deleteData.NewRow();
                        dataFill(dt, dtRow, dt1Index, 0);
                        deleteData.Rows.Add(dtRow);
                        dt1Index++;
                    }
                }
            }
            
            Console.WriteLine("NewData:");
            foreach (DataRow row1 in NewData.Rows)
            {
                foreach (string str1 in row1.ItemArray)
                {
                    Console.Write(str1);
                }
                Console.Write("\n");
            }
            Console.WriteLine("UpdateData:");
            foreach (DataRow row1 in updateData.Rows)
            {
                foreach (string str1 in row1.ItemArray)
                {
                    Console.Write(str1);
                }
                Console.Write("\n");
            }
            Console.WriteLine("DeleteData:");
            foreach (DataRow row1 in deleteData.Rows)
            {
                foreach (string str1 in row1.ItemArray)
                {
                    Console.Write(str1);
                }
                Console.Write("\n");
            }
            createExcel(NewData, "新增資料");
            createExcel(deleteData, "刪除資料");
            createExcel(updateData, "修改資料");
        }

        public static Boolean checkDataEquals(DataTable dt1, DataTable dt2, int index1, int index2, List<int> Field)
        {
            Boolean isSame = true;

            foreach (int i in Field)
            {
                if (!dt1.Rows[index1].ItemArray[i].ToString().Equals(dt2.Rows[index2].ItemArray[i].ToString()))
                {
                    isSame = false;
                    break;
                }
            }
            return isSame;
        }

        public static void createExcel(DataTable dt, string fileName)
        {
            IWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)wb.CreateSheet(fileName);

            MemoryStream ms = new MemoryStream();

            XSSFRow row = null;

            //設定儲存格樣式
            XSSFCell cell = null;
            XSSFCellStyle wrapStyle = (XSSFCellStyle)wb.CreateCellStyle();

            XSSFFont font1 = (XSSFFont)wb.CreateFont();
            //字體尺寸
            font1.FontHeightInPoints = 12;

            wrapStyle.SetFont(font1);
            wrapStyle.WrapText = true;
            wrapStyle.BorderTop = BorderStyle.Thin;
            wrapStyle.BorderLeft = BorderStyle.Thin;
            wrapStyle.BorderBottom = BorderStyle.Thin;
            wrapStyle.BorderRight = BorderStyle.Thin;
            wrapStyle.VerticalAlignment = VerticalAlignment.Center;

            sheet1.PrintSetup.Landscape = true;
            sheet1.ForceFormulaRecalculation = true;

            int rowIndex = 0;
            int field = 0;
            
            row = (XSSFRow)sheet1.CreateRow(rowIndex);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                cell = (XSSFCell)row.CreateCell(i);
                cell.CellStyle = wrapStyle;  //指定樣式
                //cell.SetCellType(CellType.String);
                cell.SetCellValue(dt.Columns[i].ColumnName);
            }
            rowIndex++;


            foreach (DataRow row1 in dt.Rows)
            {
                field = 0; 
                row = (XSSFRow)sheet1.CreateRow(rowIndex);
                foreach (string value in row1.ItemArray)
                {
                    cell = (XSSFCell)row.CreateCell(field);
                    cell.CellStyle = wrapStyle;  //指定樣式
                    //cell.SetCellType(CellType.String);
                    cell.SetCellValue(value);
                    field++;
                }
                rowIndex++;
            }

            sheet1.SetColumnWidth(0, 20 * 256);
            sheet1.SetColumnWidth(1, 30 * 256);
            //產生檔案
            FileStream FS = File.Create(fileName + ".xlsx");
            wb.Write(FS);
            FS.Close();

            Console.WriteLine("Create Excel: '" + fileName + "' Success.");
        }

        private static DataRow dataFill(DataTable dt, DataRow dtRow, int itemIndex, int index)
        {
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                dtRow[dt.Columns[j].ColumnName] = dt.Rows[itemIndex + index].ItemArray[j].ToString();
            }
            return dtRow;
        }


        public static DataTable loadExcel(string FileName)
        {
            IWorkbook wk;
            using (FileStream fs = new FileStream(".\\" + FileName, FileMode.Open, FileAccess.ReadWrite))
            {
                wk = new XSSFWorkbook(fs);
            }
            XSSFSheet sheet = (XSSFSheet)wk.GetSheet("TestSheet");
            XSSFRow row = (XSSFRow)sheet.GetRow(0);
            int lastRow = sheet.LastRowNum;
            int lastCell = row.LastCellNum;
            DataTable dt = new DataTable();
            //DataRow dtRow = dt.NewRow();
            for (int i = 0; i < lastCell; i++)
            {
                string str = Object.Equals(row.GetCell(i), null) ? " " : row.GetCell(i).ToString();
                dt.Columns.Add(str, typeof(String));
                dt.Columns[str].MaxLength = 50;
                dt.Columns[str].AllowDBNull = false;
            }

            for (int i = 1; i <= lastRow; i++)
            {
                DataRow dtRow = dt.NewRow();
                row = (XSSFRow)sheet.GetRow(i);
                for (int j = 0; j < lastCell; j++)
                {
                    string str = Object.Equals(row.GetCell(j), null) ? " " : row.GetCell(j).ToString();
                    dtRow[sheet.GetRow(0).GetCell(j).ToString()] = str;
                    //Console.Write(str + " ");
                }
                dt.Rows.Add(dtRow);
                //Console.Write("\n");
            }
            /*
            foreach (DataRow dataRow in dt.Rows)
            {
                foreach (var item in dataRow.ItemArray)
                {
                    Console.Write(item);
                }
                Console.Write("\n");
            }
            */
            return dt;
        }

        public static void make0419()
        {
            string str = string.Empty;
            
            List<string> ls = getSite();
            Console.Write(ls.Count);
            str += "<td width=\"50%\" valign=\"top\"><table width=\"100%\" border=\"0\" cellpadding=\"1\" bgcolor=\"#aaaaaa\" class='search_Select_Style' id=\"Table3\">";
            for (int i=0; i<ls.Count; i++)
            {

                if (i%2==0)
                {
                    str += "<tr>\n";
                }


                str += "<td width=\"25%\" align=\"center\" valign=\"middle\" bgcolor=\"#D3E658\"><p align=\"center\">" + ls[i] + "</p></td>";
                str += "<td valign=\"top\" bgcolor=\"#FFFFFF\"><a href=\"#\"><img src=\"download.png\" width=\"64\" height=\"64\" /></a></td>";
                str += "<td valign=\"top\" bgcolor=\"#FFFFFF\"></td>\n";
                if (i%2==1)
                {
                    str += "</tr>";
                }
                
            }
            str += "</table>\n";
            str += "</td>\n";


            using (StreamWriter outputFile = new StreamWriter(@"D:\\Brian\\WriteLine.txt", true))
            {
                outputFile.WriteLine(str);
            }
        }

        public static List<string> getSite()
        {
            List<string> ls = new List<string>();
            ls.Add("主計處 ");
            ls.Add("政風處 ");
            ls.Add("財政處 ");
            ls.Add("法制處 ");
            ls.Add("秘書處 ");
            ls.Add("新聞及國際關係處 ");
            ls.Add("民族事務委員會");
            ls.Add("民政局 ");
            ls.Add("農業局 ");
            ls.Add("勞工局 ");
            ls.Add("經濟發展局 ");
            ls.Add("水利局 ");
            ls.Add("動物防疫保護處 ");
            ls.Add("臺南市市場處 ");
            ls.Add("漁港及近海管理所 ");
            ls.Add("區公所-安定區 ");
            ls.Add("區公所-安南區 ");
            ls.Add("區公所-新化區 ");
            ls.Add("區公所-學甲區 ");
            ls.Add("區公所-北門區 ");
            ls.Add("區公所-七股區 ");
            ls.Add("區公所-大內區 ");
            ls.Add("區公所-東山區 ");
            ls.Add("區公所-關廟區 ");
            ls.Add("區公所-官田區 ");
            ls.Add("區公所-後壁區 ");
            ls.Add("區公所-佳里區 ");
            ls.Add("區公所-將軍區 ");
            ls.Add("區公所-龍崎區 ");
            ls.Add("區公所-南化區 ");
            ls.Add("區公所-楠西區 ");
            ls.Add("區公所-山上區 ");
            ls.Add("區公所-北區");
            ls.Add("區公所-下營區");
            ls.Add("區公所-新營區 ");
            ls.Add("區公所-鹽水區 ");
            ls.Add("區公所-左鎮區");
            ls.Add("戶政事務所-安平戶政事務所 ");
            ls.Add("戶政事務所-白河戶政事務所 ");
            ls.Add("戶政事務所-官田戶政事務所 ");
            ls.Add("戶政事務所-歸仁戶政事務所 ");
            ls.Add("戶政事務所-佳里戶政事務所 ");
            ls.Add("戶政事務所-仁德戶政事務所 ");
            ls.Add("戶政事務所-善化戶政事務所 ");
            ls.Add("戶政事務所-新化戶政事務所 ");
            ls.Add("戶政事務所-新營戶政事務所 ");
            ls.Add("戶政事務所-學甲戶政事務所 ");
            ls.Add("戶政事務所-安南戶政事務所 ");
            ls.Add("戶政事務所-府東戶政事務所 ");
            ls.Add("戶政事務所-麻豆戶政事務所 ");
            ls.Add("戶政事務所-玉井戶政事務所 ");
            ls.Add("戶政事務所-永康戶政事務所 ");
            ls.Add("衛生所-安南區 ");
            ls.Add("衛生所-安平區 ");
            ls.Add("衛生所-東區 ");
            ls.Add("衛生所-北區 ");
            ls.Add("衛生所-南區 ");
            ls.Add("衛生所-中西區 ");
            ls.Add("衛生所-安定區 ");
            ls.Add("衛生所-將軍區 ");
            ls.Add("衛生所-七股區 ");
            ls.Add("衛生所-佳里區 ");
            ls.Add("衛生所-學甲區 ");
            ls.Add("衛生所-新化區 ");
            ls.Add("衛生所-西港區 ");
            ls.Add("衛生所-後壁區 ");
            ls.Add("衛生所-新市區 ");
            ls.Add("衛生所-下營區 ");
            ls.Add("衛生所-仁德區 ");
            ls.Add("衛生所-歸仁區 ");
            ls.Add("衛生所-關廟區 ");
            ls.Add("衛生所-官田區 ");
            ls.Add("衛生所-六甲區 ");
            ls.Add("衛生所-柳營區 ");
            ls.Add("衛生所-麻豆區 ");
            ls.Add("衛生所-楠西區 ");
            ls.Add("衛生所-南化區 ");
            ls.Add("衛生所-白河區 ");
            ls.Add("衛生所-北門區 ");
            ls.Add("衛生所-善化區 ");
            ls.Add("衛生所-山上區 ");
            ls.Add("衛生所-新營區 ");
            ls.Add("衛生所-左鎮區 ");
            ls.Add("衛生所-大內區 ");
            ls.Add("衛生所-東山區 ");
            ls.Add("衛生所-玉井區 ");
            ls.Add("衛生所-永康區 ");
            ls.Add("衛生所-鹽水區 ");

            return ls;
        }


        public static void NPOIToExcel(List<List<List<string>>> ls, string siteName)
        {

            IWorkbook wb = new XSSFWorkbook();
            XSSFSheet sheet1 = (XSSFSheet)wb.CreateSheet(siteName);

            MemoryStream ms = new MemoryStream();

            XSSFRow row = null;

            //設定儲存格樣式
            XSSFCell cell = null;
            XSSFCellStyle wrapStyle = (XSSFCellStyle)wb.CreateCellStyle();

            XSSFFont font1 = (XSSFFont)wb.CreateFont();
            //字體尺寸
            font1.FontHeightInPoints = 12;

            wrapStyle.SetFont(font1);
            wrapStyle.WrapText = true;
            wrapStyle.BorderTop = BorderStyle.Thin;
            wrapStyle.BorderLeft = BorderStyle.Thin;
            wrapStyle.BorderBottom = BorderStyle.Thin;
            wrapStyle.BorderRight = BorderStyle.Thin;
            wrapStyle.VerticalAlignment = VerticalAlignment.Center;

            sheet1.PrintSetup.Landscape = true;
            sheet1.ForceFormulaRecalculation = true;

            int rowIndex = 0;
            int field = 0;
            int rowPartIndex = 0;

            row = (XSSFRow)sheet1.CreateRow(rowIndex);
            cell = (XSSFCell)row.CreateCell(0);
            cell.CellStyle = wrapStyle;  //指定樣式
            cell.SetCellType(CellType.String);
            cell.SetCellValue("頁面名稱");
            cell = (XSSFCell)row.CreateCell(1);
            cell.CellStyle = wrapStyle;  //指定樣式
            cell.SetCellType(CellType.String);
            cell.SetCellValue("連結");
            rowIndex++;
            rowPartIndex++;

            foreach (List<List<string>> part in ls)
            {
                int size = part.First().Count + rowPartIndex;
                for (int i = rowPartIndex; i < size; i++)
                {
                    row = (XSSFRow)sheet1.CreateRow(i);
                }

                field = 0;

                foreach (List<string> part1 in part)
                {
                    rowIndex = rowPartIndex;
                    foreach (string str in part1)
                    {
                        row = (XSSFRow)sheet1.GetRow(rowIndex);
                        cell = (XSSFCell)row.CreateCell(field);
                        cell.CellStyle = wrapStyle;  //指定樣式
                        //cell.SetCellType(CellType.String);
                        cell.SetCellValue(str);
                        rowIndex++;
                    }
                    field++;
                }
                rowPartIndex = rowIndex;
            }

            sheet1.SetColumnWidth(0, 20 * 256);
            sheet1.SetColumnWidth(1, 30 * 256);
            //產生檔案
            FileStream FS = File.Create("D:\\Brian\\" + siteName + ".xlsx");
            wb.Write(FS);
            FS.Close();
        }

        static void getSitemap(string url, string topQuery, string contentQuery, string footQuery)
        {
            if (url.Equals(""))
            {
                Console.Write("請輸入該網站的網站導覽網址:");
                url = Console.ReadLine();
            }
            string htmlContent = GetContent(url);
            List<List<List<string>>> ls = new List<List<List<string>>>();
            //foreach (string item in GetHtmlBySelector("div form a", htmlContent))
            //{
            //    Console.WriteLine(item);
            //}
            List<string> topText = WebCrawler(url, topQuery, "value");
            List<string> topLink = WebCrawler(url, topQuery, "href");
            //foreach (string item in topText)
            //{
            //    Console.WriteLine(item);
            //}
            //foreach (string item in topLink)
            //{
            //    Console.WriteLine(item);
            //}
            List<List<string>> top = new List<List<string>>();
            top.Add(topText);
            top.Add(topLink);
            //foreach (string item in GetHtmlBySelector("ul li a", htmlContent))
            //{
            //    Console.WriteLine(item);
            //}
            List<string> contentText = WebCrawler(url, contentQuery, "value");
            List<string> contentLink = WebCrawler(url, contentQuery, "href");
            //foreach (string item in contentText)
            //{
            //    Console.WriteLine(item);
            //}
            //foreach (string item in contentLink)
            //{
            //    Console.WriteLine(item);
            //}
            List<List<string>> content = new List<List<string>>();
            content.Add(contentText);
            content.Add(contentLink);

            //foreach (string item in GetHtmlBySelector("ul li a", htmlContent))
            //{
            //    Console.WriteLine(item);
            //}
            List<string> footText = WebCrawler(url, footQuery, "value");
            List<string> footLink = WebCrawler(url, footQuery, "href");
            //foreach (string item in contentText)
            //{
            //    Console.WriteLine(item);
            //}
            //foreach (string item in contentLink)
            //{
            //    Console.WriteLine(item);
            //}
            List<List<string>> foot = new List<List<string>>();
            foot.Add(footText);
            foot.Add(footLink);

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
            /*
            List<string> footAlt = new List<string>();
            List<string> footLink = new List<string>();
            foreach (string item in GetHtmlBySelector("a", htmlContent))
            {
                foreach (string str in GetHtmlBySelector("a img", htmlContent))
                {
                    if (item.Contains(str))
                    {
                        //Console.WriteLine(getAttribute(item, "href"));
                        footLink.Add(getAttribute(item, "href"));
                        break;
                    }
                }
            }
            foreach (string item in GetHtmlBySelector("a img", htmlContent))
            {
                //Console.WriteLine(getAttribute(item, "alt"));
                footAlt.Add(getAttribute(item, "alt"));
            }
            List<List<string>> foot = new List<List<string>>();
            foot.Add(footAlt);
            foot.Add(footLink);
            */
            ls.Add(top);
            ls.Add(content);
            ls.Add(foot);

            

            //foreach (string item in GetHtmlBySelector("title", htmlContent))
            //{
            //    Console.WriteLine(item);
            //}
            if (topText.Count == 0 && contentText.Count == 0 && footText.Count == 0)
            {
                Console.WriteLine("無資料，網址" + url + "可能為錯誤頁面。");
            }
            else
            {
                string siteName = getAttribute(GetHtmlBySelector("title", htmlContent).Where(o => o == o).First(), "value");
                //Console.WriteLine("topText.Count = " + topText.Count + ", contentText.Count = " + contentText.Count + ", footAlt.Count = " + footAlt.Count);
                Console.WriteLine("成功爬到 " + siteName + " 的資料，開始產生EXCEL。");
                NPOIToExcel(ls, siteName);
            }
            
        }

        static void getSitemapForTainan(string url, string topQuery, string contentQuery, string footQuery)
        {
            if (url.Equals(""))
            {
                Console.Write("請輸入該網站的網站導覽網址:");
                url = Console.ReadLine();
            }
            string htmlContent = GetContent(url);
            List<List<List<string>>> ls = new List<List<List<string>>>();

            List<string> topText = WebCrawler(url, topQuery, "value");
            List<string> topLink = WebCrawler(url, topQuery, "href");
            List<List<string>> top = new List<List<string>>();
            top.Add(topText);
            top.Add(topLink);

            List<string> contentText = WebCrawler(url, contentQuery, "value");
            List<string> contentLink = WebCrawler(url, contentQuery, "href");
            List<List<string>> content = new List<List<string>>();
            content.Add(contentText);
            content.Add(contentLink);

            List<string> footAlt = new List<string>();
            List<string> footLink = new List<string>();
            foreach (string item in GetHtmlBySelector("a", htmlContent))
            {
                foreach (string str in GetHtmlBySelector("a img", htmlContent))
                {
                    if (item.Contains(str))
                    {
                        footLink.Add(getAttribute(item, "href"));
                        break;
                    }
                }
            }
            foreach (string item in GetHtmlBySelector("a img", htmlContent))
            {
                footAlt.Add(getAttribute(item, "alt"));
            }
            List<List<string>> foot = new List<List<string>>();
            foot.Add(footAlt);
            foot.Add(footLink);
            
            ls.Add(top);
            ls.Add(content);
            ls.Add(foot);

            if (topText.Count == 0 && contentText.Count == 0 && footAlt.Count == 0)
            {
                Console.WriteLine("無資料，網址" + url + "可能為錯誤頁面。");
                using(StreamWriter outputFile = new StreamWriter(@"D:\\Brian\\WriteLine.txt", true))
                {
                    outputFile.WriteLine("無資料，網址" + url + "可能為錯誤頁面。");
                }
            }
            else
            {
                string siteName = getAttribute(GetHtmlBySelector("title", htmlContent).Where(o => o == o).First(), "value");
                Console.WriteLine("成功爬到 " + siteName + " 的資料，開始產生EXCEL。");
                NPOIToExcel(ls, siteName);
            }

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

        public static string getAttribute(string result, string attr)
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
