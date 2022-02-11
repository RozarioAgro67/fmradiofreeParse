using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;

namespace fmradiofreeParse
{
    class Program
    {
        static void Main(string[] args)
        {
            string pageUrl = "https://www.fmradiofree.com/?page=";
            List<Card> result = new List<Card>();
            using (HttpClientHandler clientHandler = new HttpClientHandler { AllowAutoRedirect = false, AutomaticDecompression = System.Net.DecompressionMethods.Deflate | System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.None })
            {
                using (HttpClient client = new HttpClient(clientHandler))
                {
                    for (int i = 1; i < 334; i++)
                    {
                        string namberPage = i.ToString();
                        var href = ParsingHref(client, pageUrl + namberPage);
                        for (int a = 0; a < href.Count; a++)
                        {
                            result.Add(ParsingCard(client, href[a]));
                        }
                        Console.WriteLine($"Спарсили страницу: {i}");
                    }

                    using (ExcelHelper excel = new ExcelHelper())
                    {
                        excel.Open(filePath: Path.Combine(Environment.CurrentDirectory, "fmradiofree.xlsx"));
                        for (int i = 1; i < result.Count; i++)
                        {
                            excel.Set(i, 1, result[i - 1].NameFm);
                            excel.Set(i, 2, result[i - 1].Email);
                            excel.Set(i, 3, result[i - 1].Telephone);
                            excel.Set(i, 4, result[i - 1].Adress);
                            excel.Set(i, 5, result[i - 1].Facebook);
                            excel.Set(i, 6, result[i - 1].Twitter);
                            excel.Set(i, 7, result[i - 1].Website);
                            excel.Set(i, 8, result[i - 1].Category);
                        }
                        excel.Save();
                    }

                    Console.WriteLine("Работу закончил");
                    Console.ReadLine();
                }
            }
        }

        private static List<string> ParsingHref(HttpClient client, string namberPage)
        {
            List<string> href = new List<string>();
            using (var resp = client.GetAsync(namberPage).Result)
            {
                
                HtmlDocument doc = new HtmlDocument();
                var html = resp.Content.ReadAsStringAsync().Result;
                doc.LoadHtml(html);
                var list = doc.DocumentNode.SelectNodes("//li[@id]//a[@href]");
                foreach (var item in list)
                {
                    href.Add(item.GetAttributeValue("href", null));

                }
            } 
            return href;
        }        

        private static Card ParsingCard(HttpClient client, string href)
        {
            Card card = new Card();
            using(var resp = client.GetAsync("https://www.fmradiofree.com" + href).Result)
            {
                var html = resp.Content.ReadAsStringAsync().Result;
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(html);
                card.NameFm = doc.DocumentNode.SelectSingleNode("//h1[@class='mdc-typography--display1 primary-span-color']").InnerText.Replace("\n", "").Replace("&#39;", "'").Trim();
                try
                {
                    card.Email = doc.DocumentNode.SelectSingleNode(".//p[contains(text(), 'E-mail')]//span").InnerText;
                }
                catch (Exception){card.Email = "";}

                try
                {
                    card.Telephone = doc.DocumentNode.SelectSingleNode(".//p[contains(text(), 'Telephone')]//span").InnerText;
                }
                catch (Exception) { card.Telephone = ""; }

                try
                {
                    card.Adress = doc.DocumentNode.SelectSingleNode(".//p[contains(text(), 'Address:')]//span").InnerText.Replace("\r\n", " ");
                }
                catch (Exception) { card.Adress = ""; }

                try
                {
                    card.Facebook = doc.DocumentNode.SelectSingleNode("//div[@class='social-buttons']//*[contains(@href, 'https://www.facebook.com')]").GetAttributeValue("href", null);
                }
                catch (Exception) { card.Facebook = ""; }


                try
                {
                    card.Twitter = doc.DocumentNode.SelectSingleNode("//div[@class='social-buttons']//*[contains(@href, 'https://twitter.com')]").GetAttributeValue("href", null);
                }
                catch (Exception) { card.Twitter = ""; }


                try
                {
                    card.Website = doc.DocumentNode.SelectSingleNode(".//span[contains(text(), 'Website')]/..//a").GetAttributeValue("href", null);
                }
                catch (Exception) { card.Website = ""; }


                try
                {
                    card.Category = doc.DocumentNode.SelectSingleNode("//span[contains(text(), 'Categories:')]/..").InnerText.Replace("Categories:&nbsp", "").Replace("\n", "").Trim();
                }
                catch (Exception) { card.Category = ""; }
            }
            return card;
        }
    }
}
