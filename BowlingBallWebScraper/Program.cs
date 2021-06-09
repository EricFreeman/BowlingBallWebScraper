using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using HtmlAgilityPack;
using CsvHelper;
using System.IO;
using System.Globalization;

namespace BowlingBallWebScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            var scraper = new BowlingBallWebScraper();
            scraper.Run();
        }
    }

    class BowlingBallWebScraper
    {
        private const string HOST = "https://www.bowlingball.com";
        private const string SHOPPINGPAGE = "/shop/all/bowling-balls/?ft1=a&limit=999999";

        public void Run()
        {
            var balls = LoadBalls();
            var urls = balls.Select(ball => ball.SelectSingleNode("a").Attributes["href"].Value);

            // asynchronously process urls
            var collection = new BlockingCollection<BowlingBall>();
            Parallel.ForEach(urls, url =>
            {
                try
                {
                    ProcessBall(url, collection);
                }
                catch(Exception e)
                {
                    Console.WriteLine("!! ERROR !!");
                    Console.WriteLine($"Failed to parse {url}. Skipping.");
                    Console.WriteLine(e);
                }
            });

            WriteCSV(collection);

            Console.WriteLine("All done! Press enter to close");
            Console.ReadLine();
        }

        private HtmlNodeCollection LoadBalls()
        {
            var web = new HtmlWeb();
            var doc = web.Load(HOST + SHOPPINGPAGE);
            var balls = doc.DocumentNode.SelectNodes("//div[@class='product_info_block']");

            Console.WriteLine($"Found {balls.Count} balls");

            return balls;
        }

        private void ProcessBall(string url, BlockingCollection<BowlingBall> collection)
        {
            Console.WriteLine($"Processing {url}");

            var web = new HtmlWeb();
            var doc = web.Load(HOST + url);
            var specTable = doc.DocumentNode.SelectSingleNode("//table[@class='specs_table']");

            if (specTable == null)
            {
                Console.WriteLine($"Failed to find spec table for {url}. Skipping.");
                return;
            }

            var specs = specTable.SelectNodes("tr").SelectMany(tableRow => tableRow.SelectNodes("td")).ToList();

            var ball = new BowlingBall();
            ball.Name = doc.DocumentNode.SelectSingleNode("//h1[@class='ProductNameText']").InnerText;
            ball.Price = doc.DocumentNode.SelectSingleNode("//input[@type='hidden' and @itemprop='price']")?.Attributes["content"]?.Value;
            ball.ProductID = GetSpec(specs, "Product ID");
            ball.Brand = GetSpec(specs, "Brand");
            ball.PerfectScaleHookRating = GetSpec(specs, "Perfect Scale Hook Rating");
            ball.RG = GetSpec(specs, "RG");
            ball.Finish = GetSpec(specs, "Finish");
            ball.BallColor = GetSpec(specs, "Ball Color");
            ball.LaneCondition = GetSpec(specs, "Lane Condition");
            ball.Coverstock = GetSpec(specs, "Coverstock");
            ball.BallQuality = GetSpec(specs, "Ball Quality");
            ball.BallWarranty = GetSpec(specs, "Ball Warranty");
            ball.FactoryFinish = GetSpec(specs, "Factory Finish");
            ball.BreakpointShape = GetSpec(specs, "Breakpoint Shape");
            ball.CoverstockName = GetSpec(specs, "Coverstock Name");
            ball.CoreName = GetSpec(specs, "Core Name");
            ball.Differential = GetSpec(specs, "Differential");
            ball.Durometer = GetSpec(specs, "Durometer");
            ball.FlarePotential = GetSpec(specs, "Flare Potential");
            ball.CoreType = GetSpec(specs, "Core Type");
            ball.Performance = GetSpec(specs, "Performance");
            ball.StormProductLine = GetSpec(specs, "Storm Product Line");
            ball.ReleaseDate = GetSpec(specs, "Release Date");
            ball.HookPotential = GetSpec(specs, "Hook Potential:");
            ball.URL = HOST + url;

            collection.Add(ball);
        }

        // spec table is series of table cells where even is spec name and odd is value
        private string GetSpec(List<HtmlNode> specs, string specToFind)
        {
            for (var i = 0; i < specs.Count; i += 2)
            {
                var specName = specs[i].InnerText;
                var specValue = specs[i + 1].InnerText;

                if (specName == specToFind)
                {
                    return specValue;
                }
            }

            return "";
        }

        private void WriteCSV(BlockingCollection<BowlingBall> balls)
        {
            var path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            Console.WriteLine($"Writing CSV to {path}\\balls.csv");

            using (var writer = new StreamWriter($"{path}\\balls.csv"))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                var headers = typeof(BowlingBall).GetFields().Select(property => property.Name).ToList();
                headers.ForEach(header => csv.WriteField(header));
                csv.NextRecord();

                foreach (var ball in balls)
                {
                    var values = headers.Select(property => ball.GetType().GetField(property).GetValue(ball)).ToList();
                    values.ForEach(value => csv.WriteField(value));
                    csv.NextRecord();
                }

                writer.Flush();
            }
        }
    }

    class BowlingBall
    {
        public string Name;
        public string ProductID;
        public string Price;
        public string Brand;
        public string PerfectScaleHookRating;
        public string RG;
        public string Finish;
        public string BallColor;
        public string LaneCondition;
        public string Coverstock;
        public string BallQuality;
        public string BallWarranty;
        public string FactoryFinish;
        public string BreakpointShape;
        public string CoverstockName;
        public string CoreName;
        public string Differential;
        public string Durometer;
        public string FlarePotential;
        public string CoreType;
        public string Performance;
        public string StormProductLine;
        public string ReleaseDate;
        public string HookPotential;
        public string URL;
    }
}
