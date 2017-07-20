using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Exchange;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Runtime;
using System.Xml;
using System.IO;
using AfterHourConsoleApplication.SDK;

namespace AfterHourConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            GetConfiguration(args);
            
            var t = new Thread(Run);
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            Console.ReadLine();
        }


        static void Run()
        {
            userLogin();

            UserConfig.getUserConfig();
            if (!UserConfig.isUserConfigValid()) {
                exit();
            }
            AfterHourImpactComputor.getAfterHourImpact();

            //generateProcessReport();
        }
        
        static void generateProcessReport() {
            Utils.writeProcessReportToFile();
            Utils.resetProcessReport();
        }

        static string getTokenFromCache()
        {
            try
            {
                string file = Utils.filePath + Utils.tokenFile;
                if (File.Exists(file))
                {
                    string[] lines = System.IO.File.ReadAllLines(file);
                    if (lines.Length > 1 && Convert.ToDateTime(lines[1]) > DateTime.Now)
                    {
                        return lines[0];
                    }
                }
            }
            catch (Exception ex)
            {
                return "";
            }
            
            return "";
        }

        static void userLogin() {

            string token = getTokenFromCache();
            if (token != "")
            {
                Utils.token = token;
                return;
            }
            string authority = ConfigurationManager.AppSettings["authority"];
            string clientID = ConfigurationManager.AppSettings["clientId"];
            string clientSecret = ConfigurationManager.AppSettings["clientSecret"];
            Uri redirectUri = new Uri(ConfigurationManager.AppSettings["redirectUri"]);
            string serverName = ConfigurationManager.AppSettings["serverName"];

            AuthenticationContext authenticationContext = new AuthenticationContext(authority);
            AuthenticationResult authenticationResult = authenticationContext.AcquireToken(serverName, clientID, redirectUri, PromptBehavior.Always);
            Utils.token = authenticationResult.AccessToken;

            writeToFileWithPath(Utils.token + "\r\n" + DateTime.Now.AddMinutes(30).ToString(), Utils.tokenFile, Utils.filePath);
            //Utils.token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6ImVuaDlCSnJWUFU1aWpWMXFqWmpWLWZMMmJjbyJ9.eyJuYW1laWQiOiJmZTkzYmZlMS03OTQ3LTQ2MGEtYTVlMC03YTU5MDZiNTEzNjNANzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3IiwidmVyIjoiRXhjaGFuZ2UuQ2FsbGJhY2suVjEiLCJhcHBjdHhzZW5kZXIiOiJodHRwczovL2FnYXZlY2RuLm8zNjV3ZXZlLWRldi5jb20vQXBwUmVhZC9Ib21lL0hvbWUtZGV2Lmh0bWwiLCJhcHBjdHgiOiJ7XCJvaWRcIjpcImI4N2FhYzEwLTQzZjgtNGNmMy1hY2FmLWVkNWUyZmE2ZmVhNFwiLFwicHVpZFwiOlwiMTAwMzdGRkU5Njc5QUYyRFwiLFwic210cFwiOlwibGltamlAbWljcm9zb2Z0LmNvbVwiLFwidXBuXCI6XCJsaW1qaUBtaWNyb3NvZnQuY29tXCIsXCJtc2V4Y2hwcm90XCI6XCJld3NcIn0iLCJpc3MiOiIwMDAwMDAwMi0wMDAwLTBmZjEtY2UwMC0wMDAwMDAwMDAwMDBANzJmOTg4YmYtODZmMS00MWFmLTkxYWItMmQ3Y2QwMTFkYjQ3IiwiYXVkIjoiMDAwMDAwMDItMDAwMC0wZmYxLWNlMDAtMDAwMDAwMDAwMDAwL291dGxvb2sub2ZmaWNlMzY1LmNvbUA3MmY5ODhiZi04NmYxLTQxYWYtOTFhYi0yZDdjZDAxMWRiNDciLCJleHAiOjE1MDAwNTk0NzksIm5iZiI6MTUwMDAzMDY3OX0.aZyTkR0EJ6ynBAlok__wHNTgq9hryY4JNnHIBfAfi5vBo8A648Jwbd6KINr_fKX5b3Xvy5sP78rM4tdr0ULHgxBQJtlO07kFF8VDEz72BlP5XR2NoqTLEbAF_eMUrLPNRMeXLawmLfQ2UQNekGUJBZgkuI5kDsBUL8PnHcsT2sKP162eLjH8I-TZgo4_oly86GwmTMG2vO0h1Mee8XkdjyrZpna4y_Ixemz7pmhJ63bO1NC2eQZkB0NPv6pk3aawz76L6O5lwzuKgIAGolnP1B76ZlsbDiIOClFsKImszRctZROgTc0dM7d1I1DxUuQiRDiczGjEVpAayLBK4Ixx4w";
        }

        public static void writeToFileWithPath(string content, string fileName, string path)
        {
            if (!Directory.Exists(path))
            {
                DirectoryInfo di = Directory.CreateDirectory(path);
            }

            FileStream fileStream = new FileStream(path + fileName, FileMode.Create, FileAccess.Write);
            StreamWriter sWriter = new StreamWriter(fileStream);
            sWriter.Write(content);
            sWriter.Close();
        }

        static void Help()
        {
            Console.WriteLine("Please input valid parameters...\n");

            Console.WriteLine("Example: AfterHourConsoleApplication -days 7 -startConv false -directReply true\n");
            Console.WriteLine("\t-days(or -d): number between 0 and 30, we'll collect your last N days' data. default 7");
            Console.WriteLine("\t-startConv(or -s): bool, you must be the starter of the conversation if true. default false");
            Console.WriteLine("\t-directReply(or -r): bool, we only count those emails direct reply to you if true. default false");
        }
        static void GetConfiguration(string[] args)
        {
            if (args.Length % 2 == 1)
            {
                exit();
            }

            for (var i = 0; i < args.Length; i = i + 2)
            {
                if (args[i].ToLower() == "-days" || args[i].ToLower() == "-d")
                {
                    try
                    {
                        var days = Convert.ToInt32(args[i + 1]);
                        if (days > 0 && days <= 60)
                        {
                            Utils.startRange = Utils.generatePriorDayBreak(days);
                            Utils.days = days;
                            Utils.lastServeralDaysText = " during last " + days.ToString() + " days.";
                        }
                    }
                    catch
                    {
                        Console.WriteLine("parameter for -days/-d is invalid. use '-days N', N is between 0 and 30.");
                        exit();
                    }
                }
                else if (args[i].ToLower() == "-startconv" || args[i].ToLower() == "-s")
                {
                    try
                    {
                        Utils.shouldStartConv = Convert.ToBoolean(args[i + 1]);
                    }
                    catch
                    {
                        Console.WriteLine("parameter for -startConv/-s is invalid. use '-startConv true' or '-startConv false'.");
                        exit();
                    }
                }
                else if (args[i].ToLower() == "-directreply" || args[i].ToLower() == "-r")
                {
                    try
                    {
                        Utils.shouldDirectReply = Convert.ToBoolean(args[i + 1]);
                    }
                    catch
                    {
                        Console.WriteLine("parameter for -directReply/-r is invalid. use '-directReply true' or '-directReply false'.");
                        exit();
                    }
                }
            }

        }

        static void exit()
        {
            Help();
            Console.ReadLine();
            System.Environment.Exit(0);
        } 
    }
}
