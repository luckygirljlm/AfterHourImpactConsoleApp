using System;
using System.Collections;
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
    static class UserConfig
    {
        private static bool _isUserConfigValid = true;
        private static SDK.WorkingHours chinaConfig = null;
        private static SDK.WorkingHours usaConfig = null;

        public static bool isUserConfigValid()
        {
            return _isUserConfigValid;
        }

        public static void getUserConfig()
        {
            string userConfigurationResponse = AfterHourImpactComputor.ExecuteRequest(new GetUserConfigurationRequest(), "UserConfiguration", false);
            Utils.userWorkingHourConfig = ParseUtils.userConfigurationPropsParsed(userConfigurationResponse);

            try
            {
                string[] oldLines = System.IO.File.ReadAllLines(Utils.configFileName);
                ArrayList lines = new ArrayList();
                for(var i=0;i< oldLines.Length;i++) {
                    if (parse(oldLines[i]) != "")
                    {
                        oldLines[i] = oldLines[i].Replace(" ", "").Replace("\t", "");
                        if (oldLines[i] != "")
                            lines.Add(oldLines[i].ToLower());
                    }
                }
                string[] newLines = new string[lines.Count];
                for (var i = 0; i < lines.Count; i++)
                    newLines[i] = lines[i].ToString();

                Dictionary<string, ArrayList> Objects = parseObject(newLines);

                if (Objects.ContainsKey("chinaconfig"))
                {
                    UserConfig.chinaConfig = parseWorkingHour(Objects["chinaconfig"], Utils.defaultWorkingHours);
                }
                if (Objects.ContainsKey("redmondconfig"))
                {
                    UserConfig.usaConfig = parseWorkingHour(Objects["redmondconfig"], Utils.defaultWorkingHours);
                }
                if (Objects.ContainsKey("chinamembers"))
                {
                    parseTeamMembers(true, Objects["chinamembers"]);
                }
                if (Objects.ContainsKey("redmondmembers"))
                {
                    parseTeamMembers(false, Objects["redmondmembers"]);
                }

                foreach (var key in Objects.Keys)
                {
                    if (key != "chinaconfig" && key != "redmondconfig" && key != "chinamembers" && key != "redmondmembers")
                    {

                        if (Utils.receipientsConfig.ContainsKey(key))
                        {
                            SDK.WorkingHours wk = parseWorkingHour(Objects[key], Utils.receipientsConfig[key]);
                            Utils.receipientsConfig[key] = wk;
                        }
                        else
                        {
                            SDK.WorkingHours wk = parseWorkingHour(Objects[key], Utils.defaultWorkingHours);
                            Utils.receipientsConfig.Add(key, wk);
                        }
                    }
                }

               // printLog();
            }
            catch (Exception ex) {
                _isUserConfigValid = false;
                Console.WriteLine(ex.ToString());
            }
        }

        static void printLog()
        {
            Console.WriteLine("shouldStartConv:" + Utils.shouldStartConv);
            Console.WriteLine("shouldDirectReply:" + Utils.shouldDirectReply);
            Console.WriteLine(Utils.receipientsConfig.Count);
            foreach (var key in Utils.receipientsConfig.Keys)
            {
                Console.WriteLine(key);
                printWorkingHour(Utils.receipientsConfig[key]);
            }
        }

        public static string printWorkingHour(SDK.WorkingHours wk)
        {
            string str = wk.StartDate.ToString() + " " + wk.EndDate.ToString() + " " + wk.TimeZoneBias.ToString() + " ";
            //Console.WriteLine(wk.StartDate);
            //Console.WriteLine(wk.EndDate);
            //Console.WriteLine(wk.TimeZoneBias);
            foreach (var i in wk.WorkDays)
            {
                //Console.Write(i);
                str = str + i.ToString();
            }
            str = str + "\r\n\r\n";
            //Console.WriteLine("--------------");
            return str;
        }

        static void parseTeamMembers(bool isChina, ArrayList nameList)
        {
            for (var i = 0; i < nameList.Count; i++)
            {
                Utils.receipientsConfig[nameList[i].ToString()] = isChina ? UserConfig.chinaConfig : UserConfig.usaConfig;
            }
        }

        static SDK.WorkingHours parseWorkingHour(ArrayList values, SDK.WorkingHours defaultWk)
        {
            SDK.WorkingHours result = new SDK.WorkingHours
            {
                StartDate = defaultWk.StartDate,
                EndDate = defaultWk.EndDate,
                TimeZoneBias = defaultWk.TimeZoneBias,
                WorkDays = defaultWk.WorkDays
            };
            bool isValid = true;
            for (int i = 0; i < values.Count; i++)
            {
                string[] pairs = values[i].ToString().Split(new char[] { ':' }, 2);
                isValid = true;
                if (pairs.Length != 2)
                {
                    isValid = false;
                    break;
                }
                else
                {
                    string key = parse(pairs[0]);
                    string value = parse(pairs[1]).Trim('"');
                    if (key == "start")
                    {
                        isValid = verifyStartAndEndTime(value);
                        if (isValid)
                        {
                            result.StartDate = value;
                        }
                    }
                    else if (key == "end")
                    {
                        isValid = verifyStartAndEndTime(value);
                        if (isValid)
                        {
                            result.EndDate = value;
                        }
                    }
                    else if (key == "timezone")
                    {
                        result.TimeZoneBias = Convert.ToInt32(value.Trim());
                    }
                    else if (key == "workdays")
                    {
                        int[] workdays = new int[value.Length];
                        for (var j = 0; j < value.Length; j++)
                        {
                            if (value[j] != '0' && value[j] != '1')
                            {
                                isValid = false;
                                break;
                            }
                            workdays[j] = value[j] == '0' ? 0 : 1;
                        }
                        if (value.Length != 7)
                            isValid = false;
                        if (isValid)
                            result.WorkDays = workdays;
                    }
                    else {
                        isValid = false;
                    }

                    if (!isValid) {
                        break;
                    }
                }

            }
            if (!isValid) {
                _isUserConfigValid = false;
                Console.WriteLine("Error when parsing WorkingHour in config file...Please check the file...\n\n");
                Console.WriteLine("Here is the error setting:");
                foreach (var text in values) {
                    Console.WriteLine(text);
                }
                Console.WriteLine("------------");
            }
            return result;
        }

        static bool verifyStartAndEndTime(string time)
        {
            string[] tmp = time.Split(new char[] { ':' });
            if (tmp.Length != 3)
                return false;
            if (tmp[0].Length != 2 || !(Convert.ToInt32(tmp[0]) >= 0 && Convert.ToInt32(tmp[0]) <= 23))
                return false;
            if (tmp[1].Length != 2 || !(Convert.ToInt32(tmp[1]) >= 0 && Convert.ToInt32(tmp[1]) <= 59))
                return false;
            if (tmp[2].Length != 2 || !(Convert.ToInt32(tmp[2]) >= 0 && Convert.ToInt32(tmp[2]) <= 59))
                return false;
            return true;
        }

        static Dictionary<string, ArrayList> parseObject(string[] lines)
        {
            Dictionary<string, ArrayList> Objects = new Dictionary<string, ArrayList>();
            for (int i = 0; i < lines.Length; i++)
            {
                string line = parse(lines[i]);
                if (line != "")
                {
                    string key = line.Split('{')[0];

                    if (Objects.ContainsKey(key))
                    {
                        Objects[key] = new ArrayList();
                    }
                    else
                    {
                        Objects.Add(key, new ArrayList());
                    }
                    if (line.Split('{').Length > 1) {
                        string[] values = line.Split('{')[1].Split(',');
                        foreach(var value in values) {
                            string parsedValue = parse(value);
                            if(parsedValue!="") 
                                Objects[key].Add(value);
                        }
                    }

                    for (i++; i < lines.Length; i++)
                    {
                        line = parse(lines[i]);
                        if (line.Contains("}"))
                            break;
                        string value = "";
                        if (line == "" || line == "{")
                        {
                            continue;
                        }
                        else if (line[0] == '{')
                        {
                            value = line.TrimStart(new char[] { '{' }).TrimEnd(new char[] { ',' });
                        }
                        else
                        {
                            value = line.TrimEnd(new char[] { ',' });
                        }
                        Objects[key].Add(value);
                    }

                    if (line.Contains("}") && line != "}")  // last line
                    {
                        Objects[key].Add(line.TrimEnd(new char[] { '}' }).TrimEnd(new char[] { ',' }));
                    }
                }
            }
            return Objects;
        }

        static string parse(string line)
        {
            line = line.ToLower().Trim();
            if (!string.IsNullOrEmpty(line) && !string.IsNullOrWhiteSpace(line) && line[0] != '#')
            {
                return line;
            }
            else
            {
                return "";
            }
        }
    }
}
