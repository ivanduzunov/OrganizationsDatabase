using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using startUp.Models;
using HtmlAgilityPack;
using startUp;
using Microsoft.Office.Interop.Excel;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ImportHtml
{
    public class Import
    {
        public static void Main(string[] args)
        {
            var dir = @"C:\Users\Mihail\Documents\visual studio 2017\Projects\OrganizationsDatabase\CentralRegisterOfNonprofitOrganizations\startUp\Models\List.xls";

            var context = new OrganizationsContext();
            using (context)
            {
                var link = context.Organizations.Select(o => o.Link).ToList().First();
                string obj = string.Empty;
                string ways = string.Empty;
                string subjOfActivity = string.Empty;
                WebClient wc = new WebClient();
                byte[] webData = wc.DownloadData(link);
                var text = Encoding.UTF8.GetString(webData);

                var pattern = @"<h4>Предмет на дейност<\/h4>\n\s*<.*>\n\s*<.*>\n\s*<.*>([^%]+)<\/pre>";
                MatchCollection collection = Regex.Matches(text, pattern);

                foreach (Match match in collection)
                {
                    Console.WriteLine(match.Groups[1].Value);
                }


            }

        }
        public static void AddOrganization(Organization org)
        {
            var context = new OrganizationsContext();

            using (context)
            {
                context.Organizations.Add(org);
                context.SaveChanges();
            }
            Console.WriteLine($"Organization {org.Name} added to the database!");

        }
        public static void ReadFromExcel(string directory)
        {

            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(directory);
            Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;
            Console.WriteLine("Press any key to start importing organizations.");
            Console.ReadKey();

            for (int i = 2; i < 50; i++)
            {
                var orgName = xlRange.Cells[i, 1].Value2.ToString();
                var type = xlRange.Cells[i, 2].Value2.ToString();
                var area = string.Empty;
                var county = string.Empty;
                var city = string.Empty;

                try
                {
                    area = xlRange.Cells[i, 3].Value2.ToString();
                }
                catch (Exception)
                {
                    area = "";
                }
                try
                {
                    county = xlRange.Cells[i, 4].Value2.ToString();
                }
                catch (Exception)
                {
                    county = "";
                }
                try
                {
                    city = xlRange.Cells[i, 5].Value2.ToString();
                }
                catch (Exception)
                {
                    city = "";
                }

                var hyperlinkAdress = xlRange.Cells[i, 1].Hyperlinks(1).Address;
                string objectives = string.Empty;
                string ways = string.Empty;
                string subjectOfActivity = string.Empty;
                //Problem
                ExtractInfoFromHyperlink(hyperlinkAdress, objectives, ways, subjectOfActivity);

                Organization org = new Organization()
                {
                    Name = orgName,
                    Type = type,
                    Area = area,
                    County = county,
                    City = city,
                    Objectives = objectives,
                    Ways = ways,
                    SubjectOfActivity = subjectOfActivity,
                    Link = hyperlinkAdress
                };

                AddOrganization(org);
            }
        }
        public static void ExtractInfoFromHyperlink(string adress, string obj, string ways, string subjOfActivity)
        {
            WebClient wc = new WebClient();
            byte[] webData = wc.DownloadData(adress);
            var text = Encoding.UTF8.GetString(webData);
            var objectivesPattern = @"<h4>Цели<\/h4>\n\s*<.*>\n\s*<.*>\n\s*<.*>([А-Яа-я\s:;.,\w\\'""-=]*)<\/pre>";
            var waysPattern = @"<h4>Средства<\/h4>\n\s*<.*>\n\s*<.*>\n\s*<.*>([А-Яа-я\s:;.,\w\\""-=]*)<\/pre>";
            var subjectOfActivityPattern = @"<h4>Предмет на дейност<\/h4>\n\s*<.*>\n\s*<.*>\n\s*<.*>([А-Яа-я\s:;.,\w\\'""]*)<\/pre>";

            MatchCollection objectivesCollection = Regex.Matches(text, objectivesPattern);
            MatchCollection waysCollection = Regex.Matches(text, waysPattern);
            MatchCollection subjectOfActivityCollection = Regex.Matches(text, subjectOfActivityPattern);

            if (objectivesCollection.Count > 0)
            {
                obj = objectivesCollection[0].Groups[1].Value.ToString();
            }
            else
            {
                obj = "";
            }
            if (waysCollection.Count > 0)
            {
                ways = waysCollection[0].Groups[1].Value.ToString();
            }
            else
            {
                ways = "";
            }
            if (subjectOfActivityCollection.Count > 0)
            {
                subjOfActivity = subjectOfActivityCollection[0].Groups[1].Value.ToString();
            }
            else
            {
                subjOfActivity = "";
            }
        }

    }

}


