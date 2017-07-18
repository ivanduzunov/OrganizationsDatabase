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
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ImportHtml
{
    public class Import
    {
        public static void Main(string[] args)
        {
            var dir = @"C:\Users\Mihail\Documents\visual studio 2017\Projects\OrganizationsDatabase\CentralRegisterOfNonprofitOrganizations\startUp\Models\List.xls";
        }

        public static void ReadFromExcel(string directory)
        {

            Application xlApp = new Application();
            Workbook xlWorkbook = xlApp.Workbooks.Open(directory);
            Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Range xlRange = xlWorksheet.UsedRange;
            Console.WriteLine("Press any key to start importing organizations.");
            Console.ReadKey();

            for (int i = 2; i < 1048; i++)
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
                List<string> tokens = ExtractInfoFromHyperlink(hyperlinkAdress);


                objectives = tokens[0];
                ways = tokens[1];
                subjectOfActivity = tokens[2];


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

                //Console.WriteLine(org.Name.ToUpper());
                //Console.WriteLine("ЦЕЛИ:" + Environment.NewLine + org.Objectives);
                //Console.WriteLine("СРЕДСТВА:" + Environment.NewLine + org.Ways);
                //Console.WriteLine("ПРЕДМЕТ НА ДЕЙНОСТ:" + Environment.NewLine + org.SubjectOfActivity);
                //Console.WriteLine();


                AddOrganization(org);
            }
            Console.WriteLine("SUCCESS !!!");
        }
        public static List<string> ExtractInfoFromHyperlink(string adress)
        {
            string obj = String.Empty;
            string ways = string.Empty;
            string subjOfActivity = string.Empty;
            List<string> toReturn = new List<string>();
            WebClient wc = new WebClient();
            byte[] webData = wc.DownloadData(adress);
            var text = Encoding.UTF8.GetString(webData);
            var objectivesPattern = @"<h4>Цели<\/h4>\n*\s*<.*>\n*\s*<.*>\n*\s*<.*>([^%]+)<\/pre>";
            var waysPattern = @"<h4>Средства<\/h4>\n*\s*<.*>\n*\s*<.*>\n*\s*<.*>([^%]+)<\/pre>";
            var subjectOfActivityPattern = @"<h4>Предмет на дейност<\/h4>\n*\s*<.*>\n*\s*<.*>\n*\s*<.*>([^%]+)<\/pre>";
            
            MatchCollection objectivesCollection = Regex.Matches(text, objectivesPattern);
            MatchCollection waysCollection = Regex.Matches(text, waysPattern);
            MatchCollection subjectOfActivityCollection = Regex.Matches(text, subjectOfActivityPattern);

            if (objectivesCollection.Count > 0)
            {
                List<string> objFull = objectivesCollection[0].Groups[1].Value.Split(new string[] { "</div>" }, StringSplitOptions.None).ToList();
                obj = objFull[0].Trim();
            }
            else
            {
                obj = "";
            }
            if (waysCollection.Count > 0)
            {
                var waysFull = waysCollection[0].Groups[1].Value.Split(new string[] { "</div>" }, StringSplitOptions.None).ToList();
                ways = waysFull[0].Trim();
            }
            else
            {
                ways = "";
            }
            if (subjectOfActivityCollection.Count > 0)
            {
                var subjOfActivityFull = subjectOfActivityCollection[0].Groups[1].Value.Split(new string[] { "su</div>" }, StringSplitOptions.None).ToList();
                subjOfActivity = subjOfActivityFull[0].Trim();
            }
            else
            {
                subjOfActivity = "";
            }
            toReturn.Add(obj);
            toReturn.Add(ways);
            toReturn.Add(subjOfActivity);
            return toReturn;
        }
        public static void AddOrganization(Organization org)
        {
            var context = new OrganizationsContext();

            using (context)
            {
                if (context.Organizations.Where(o => o.Name == org.Name).FirstOrDefault() == null)
                {
                    context.Organizations.Add(org);
                    context.SaveChanges();
                    Console.WriteLine($"{org.Name.ToUpper()} added to the database successfully!"); 
                }
                else
                {
                    Console.WriteLine($"Organization {org.Name} is already in the Organization database! Org. NOT added.");
                }
            }


        }
    }
}


