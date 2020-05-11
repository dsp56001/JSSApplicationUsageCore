using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using JSSUsage.Computer_Application_Usage;
using JSSUsage.JSSComputersGroupsByName;
using Newtonsoft.Json;
using System.Net.Http;
using System.Net.Http.Headers;
using CsvHelper;
using Microsoft.Extensions.Configuration;
using System.Globalization;

namespace JSSComputerUsageCore3
{
    public class Program
    {

        static string CONF_jssServer, CONF_groupName, CONF_outFile, CONF_startDate, CONF_endDate;
        static string CONF_jssUsername, CONF_jssPassword;


        public static void Main(string[] args)
        {

            DateTime startDate, endDate;
            endDate = System.DateTime.Now;
            startDate = endDate.AddYears(-1);

            var projConfig = new ConfigurationBuilder();

            projConfig.SetBasePath(Directory.GetCurrentDirectory());
            projConfig.AddJsonFile("settings.json");

            var config = projConfig.Build();

            CONF_jssServer = config["settings:jssServer"];
            CONF_groupName = config["settings:groupName"];
            CONF_outFile = config["settings:outFile"];
            CONF_startDate = startDate.ToString("d"); //2016-01-01
            CONF_endDate = endDate.ToString("d");
            CONF_jssUsername = config["settings:jssUsername"];
            CONF_jssPassword = config["settings:jssPassword"];

            foreach (string arg in args)
            {
                switch (arg)
                {
                    case "-help":
                    case "-h":
                        ConsoleHelp();
                        return;

                    case "-jssServer":
                        CONF_jssServer = args[Array.IndexOf(args, "-jssServer") + 1];
                        break;
                    case "-jssUsername":
                        CONF_jssUsername = args[Array.IndexOf(args, "-jssUsername") + 1];
                        break;
                    case "-jssPassword":
                        CONF_jssPassword = args[Array.IndexOf(args, "-jssPassword") + 1];
                        break;
                    case "-groupName":
                        CONF_groupName = args[Array.IndexOf(args, "-groupName") + 1];
                        break;
                    case "-outFile":
                        CONF_outFile = args[Array.IndexOf(args, "-outFile") + 1];
                        break;
                    case "-startDate":
                        string strStartDate = args[Array.IndexOf(args, "-startDate") + 1];
                        DateTime.TryParse(strStartDate, out startDate);
                        CONF_startDate = startDate.ToString("d");

                        break;
                    case "-endDate":
                        string strEndDate = args[Array.IndexOf(args, "-endDate") + 1];
                        DateTime.TryParse(strEndDate, out endDate);


                        CONF_endDate = endDate.ToString("d");
                        break;
                }
            }


            Task.Run(async () =>
            {
                // Do any async anything you need here without worry
                await Run();
            }).Wait();


            Console.ReadKey();
        }

        public static async Task Run()
        {


            Console.WriteLine("JSSUsage -help for help");
            Console.Write(string.Format("jssServer : {0} \tgroupName : {1}", CONF_jssServer, CONF_groupName));
            Console.WriteLine(string.Format("\toutFile : {0}", CONF_outFile));
            Console.WriteLine(string.Format("Start Date : {0} \tEndDate : {1}", GetJSSDate(CONF_startDate), GetJSSDate(CONF_endDate)));

            var strDone = await UpdateWriteCVS();
            Console.WriteLine(strDone);
        }

        private static void ConsoleHelp()
        {
            System.Console.WriteLine("JSSUsage outputs a csv file from a JSS server showing computer uages for a JSS groups");
            System.Console.WriteLine("Usage: -jssServer serverName [sma-jss.iam.local:8443]");
            System.Console.WriteLine("Usage: -groupName outpPutFileName [All CMI]");
            System.Console.WriteLine("Usage: -outFile outPutFileName [outfile.csv]");
            System.Console.WriteLine("Usage: -startDate []");
            System.Console.WriteLine("Usage: -endDate []");
        }
        static CSVComputerUsageRecord record;
        static async Task<string> UpdateWriteCVS()
        {
            record = new CSVComputerUsageRecord();
            List<CSVComputerUsageRecord> records = new List<CSVComputerUsageRecord>();

            JSSUsage.JSSComputersGroupsByName.Rootobject GroupsRequest = await GetComputersByGroupName(CONF_groupName);
            if (GroupsRequest == null)
            {


                return string.Format("group not found");
            }
            foreach (JSSUsage.JSSComputersGroupsByName.Computer c in GroupsRequest.computer_group.computers)
            {
                var computerUsage = await GetComputerUsageByName(c.name);
                if (computerUsage != null)
                {
                    foreach (JSSUsage.Computer_Application_Usage.Computer_Application_Usage cau in computerUsage.computer_application_usage)
                    {
                        if (cau.apps != null)
                        {
                            foreach (JSSUsage.Computer_Application_Usage.App app in cau.apps)
                            {
                                record = new CSVComputerUsageRecord()
                                {
                                    appName = app.name,
                                    computerId = c.id,
                                    computerName = c.name,
                                    date = cau.date,
                                    foreground = app.foreground,
                                    open = app.open,
                                    version = app.version
                                };
                                //csv.WriteRecord(record);
                                //csv.Flush();
                                records.Add(record);
                            }
                        }
                    }
                }

                Console.WriteLine(c.name);
            }

            bool fileExists = System.IO.File.Exists(CONF_outFile);

            using (StreamWriter writer = File.CreateText(CONF_outFile))
            {
                
                using (CsvWriter csv = new CsvWriter(writer, CultureInfo.CurrentCulture, true))
                {
                    //csv.Configuration.AutoMap<CSVComputerUsageRecord>();
                    
                    if (!fileExists)
                    {
                        csv.WriteHeader<CSVComputerUsageRecord>();
                    }
                    csv.WriteRecords<CSVComputerUsageRecord>(records);
                    
                }
            }
            return string.Format("Done Press any key to exit.");
        }

        async static Task<JSSUsage.JSSComputersGroupsByName.Rootobject> GetComputersByGroupName(string GroupName)
        {
            string URL = string.Format("https://{1}/JSSResource/computergroups/name/{0}",
               WebUtility.UrlEncode(GroupName), CONF_jssServer);
            string DATA = "";
            string JSON = await WebRequestinJson2(URL, DATA);

            JSSUsage.JSSComputersGroupsByName.Rootobject cgbn = JsonConvert.DeserializeObject<JSSUsage.JSSComputersGroupsByName.Rootobject>(JSON);

            return cgbn;
        }

        static async Task<JSSUsage.Computer_Application_Usage.Rootobject> GetComputerUsageByName(string ComputerName)
        {
            string URL = string.Format("https://{1}/JSSResource/computerapplicationusage/name/{0}/{2}_{3}",
               ComputerName, CONF_jssServer, GetJSSDate(CONF_startDate), GetJSSDate(CONF_endDate));

            string DATA = "";
            string JSON = await WebRequestinJson2(URL, DATA);

            JSSUsage.Computer_Application_Usage.Rootobject cau = JsonConvert.DeserializeObject<JSSUsage.Computer_Application_Usage.Rootobject>(JSON);

            return cau;
        }

        /// <summary>
        /// JSS API expects dates with yyy-MM-dd instead of /
        /// </summary>
        /// <param name="strDate"></param>
        /// <returns></returns>
        public static string GetJSSDate(string strDate)
        {
            DateTime dt = DateTime.Parse(strDate);
            return ((DateTime)dt).ToString(@"yyyy-MM-dd");
        }

        static async Task<string> WebRequestinJson2(string url, string postData)
        {
            Uri uri = new Uri(url);

            string ret = String.Empty;
            try
            {

                var credentials = new NetworkCredential(CONF_jssUsername, CONF_jssPassword);
                using (var handlerCreds = new HttpClientHandler { Credentials = credentials })
                using (var client = new HttpClient(handlerCreds))
                {
                    try
                    {
                        client.BaseAddress = uri;

                        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                        var response = await client.GetAsync(uri);
                        response.EnsureSuccessStatusCode(); // Throw in not success

                        var stringResponse = await response.Content.ReadAsStringAsync();
                        ret = stringResponse;

                        //Console.WriteLine($"First post is {JsonConvert.SerializeObject(posts.First())}");
                    }
                    catch (HttpRequestException e)
                    {
                        Console.WriteLine($"Request exception: {e.Message}");
                    }
                    finally
                    {
                        Console.Write("..");
                    }
                }
            }
            catch (System.Net.WebException ex)
            {
                Console.WriteLine("Exception accessing {0}", url);
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.Write("..");
            }
            return ret;
        }


    }

    [Serializable]
    class CSVComputerUsageRecord
    {
        public int computerId { get; set; }
        public string computerName { get; set; }
        public string date { get; set; }
        public string appName { get; set; }
        public string version { get; set; }
        public int foreground { get; set; }
        public int open { get; set; }
        //public string siteName { get; set; }
    }
}
