using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

using CsQuery;
using BCryptHelper = BCrypt.Net.BCrypt;

//namespace TestMakeTools
//{
class Tools
{
    public class Cryptography
    {

        public static string Hash(string Text, out string Salt)
        {
            Salt = BCryptHelper.GenerateSalt();
            return BCryptHelper.HashPassword(Text, Salt);
        }
        public static bool VerifyHash(string Text, string HashedText)
        {
            return BCryptHelper.Verify(Text, HashedText);
        }

    }

    public static string getHTMLAttribute(string result, string attr)
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
//}
