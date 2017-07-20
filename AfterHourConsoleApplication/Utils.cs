using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using AfterHourConsoleApplication.SDK;

namespace AfterHourConsoleApplication
{
    static class Utils
    {
        public static SDK.WorkingHours userWorkingHourConfig = null;  //in user's timezone
        public static Dictionary<string, SDK.WorkingHours> receipientsConfig = new Dictionary<string, SDK.WorkingHours>();
        public static bool shouldStartConv = false;
        public static bool shouldDirectReply = false;

        public static SDK.WorkingHours defaultWorkingHours = new WorkingHours
        {
            StartDate = "09:00:00",
            EndDate = "18:00:00",
            TimeZoneBias = 420,
            WorkDays = new int[] { 0, 1, 1, 1, 1, 1, 0 }
        };
        public static string token = "";
        public static string senderName = "";
        public static DateTime startRange = generatePriorDayBreak(7);
        public static int days = 7;
        public static string lastServeralDaysText = " during last 7 days.";
        public static string filePath = "./AfterHourReports/";


        public static string configFileName = "WorkingHour.config";
        public static string tokenFile = "token.txt";
        private static string processFileName = "ProcessReport.txt";
        private static string processReportText = "";

        public static string resultFileName = "ResultReport";
        public static string getFilePath() {
            if (days > 0) {
                return filePath + "Reports_" + days.ToString() + "days/";
            }
            return filePath;
        }

        public static DateTime generateCurrentDayBreak() //local
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd") + " 00:00:00";
            return Convert.ToDateTime(today);
        }

        public static void addToProcessReport(string content) {
            processReportText = processReportText + content;
        }

        public static void resetProcessReport() {
            processReportText = "";
        }

        public static void writeProcessReportToFile() {
            string file = getFilePath() + processFileName;
            FileStream fileStream = new FileStream(file, FileMode.Create, FileAccess.Write);
            StreamWriter sWriter = new StreamWriter(fileStream);
            sWriter.Write(processReportText);
            sWriter.Close();
        }

        public static DateTime generatePriorDayBreak(int numberOfDaysBackFromCurrent) //local
        {
            DateTime today = generateCurrentDayBreak();
            return today.AddDays(-numberOfDaysBackFromCurrent);
        }

        public static string toISOString(DateTime date) //date:local, return:UTC
        {
            string isoString = date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
            return isoString;
        }

        public static string getDateInSlashFormat(DateTime date) {  //date:local return UTC
            int month = date.ToUniversalTime().Month;
            int day = date.ToUniversalTime().Day;
            int year = date.ToUniversalTime().Year;

            string formatDate = convertSingleToDoubleNumber(month) + "/" + convertSingleToDoubleNumber(day) + "/" + year.ToString();
            return formatDate;
        }

        private static string convertSingleToDoubleNumber(int number) {
            return number < 10 ? ("0" + number.ToString()) : number.ToString();
        }

        private static DateTime parseWorkingHourInSpecificDay(DateTime day, string workingTime)  //date: 08:00:00
        {
            string workingHour = day.ToString("yyyy-MM-dd") + " " + workingTime;
            return Convert.ToDateTime(workingHour);
        }

        public static bool isAfterHour(DateTime date, SDK.WorkingHours workingHourConfig)  //date: UTC
        {
            DateTime localDate = date.AddMinutes(-workingHourConfig.TimeZoneBias);
            if (!isWorkDay(localDate, workingHourConfig))
                return true;

            DateTime startWorkHourToday = parseWorkingHourInSpecificDay(localDate, workingHourConfig.StartDate);
            DateTime endWorkHourToday = parseWorkingHourInSpecificDay(localDate, workingHourConfig.EndDate);

            bool isAfterHour = localDate < startWorkHourToday || localDate > endWorkHourToday;
            return isAfterHour;
        }

        public static DateTime getLastEndWorkHour(DateTime date, SDK.WorkingHours workingHourConfig) // date: UTC and ensure date is in after hour before call this, return UTC
        {
            DateTime localDate = date.AddMinutes(-workingHourConfig.TimeZoneBias);
            DateTime startHourCurrentLocalDay = parseWorkingHourInSpecificDay(localDate, workingHourConfig.StartDate);
            DateTime workDay = localDate.AddDays(0);
            if (isWorkDay(localDate, workingHourConfig) && localDate < startHourCurrentLocalDay)
            {
                workDay = localDate.AddDays(-1);
            }
            
            while (!isWorkDay(workDay, workingHourConfig))
            {
                workDay = workDay.AddDays(-1);
            }

            DateTime endWorkHour = parseWorkingHourInSpecificDay(workDay, workingHourConfig.EndDate).AddMinutes(workingHourConfig.TimeZoneBias);
            return endWorkHour;
        }

        public static DateTime getNextStartWorkHour(DateTime date, SDK.WorkingHours workingHourConfig) // date: UTC and ensure date is in after hour before call this, return UTC
        {
            DateTime localDate = date.AddMinutes(-workingHourConfig.TimeZoneBias);
            DateTime startHourCurrentLocalDay = parseWorkingHourInSpecificDay(localDate, workingHourConfig.StartDate);
            if (isWorkDay(localDate, workingHourConfig) && localDate < startHourCurrentLocalDay)
            {
                return startHourCurrentLocalDay.AddMinutes(workingHourConfig.TimeZoneBias);
            }

            DateTime nextLocalWorkDay = localDate.AddDays(1);
            while (!isWorkDay(nextLocalWorkDay, workingHourConfig))
            {
                nextLocalWorkDay = nextLocalWorkDay.AddDays(1);
            }

            DateTime startWorkHour = parseWorkingHourInSpecificDay(nextLocalWorkDay, workingHourConfig.StartDate).AddMinutes(workingHourConfig.TimeZoneBias);
            return startWorkHour;
        }

        public static bool isAfterHourReceivedAndReply(DateTime sentDate, DateTime receivedDate, SDK.WorkingHours workingHourConfig) // sent/received date in UTC
        {
            if(isAfterHour(receivedDate, workingHourConfig)) // received in after hour
            {
                DateTime checkRangeEnd = getNextStartWorkHour(receivedDate, workingHourConfig);

                return sentDate > receivedDate && sentDate < checkRangeEnd;
            }
            return false;
            
        }

        public static bool isWorkDay(DateTime date, SDK.WorkingHours workingHourConfig)    //date in local 
        {
            DayOfWeek dayInWeek = date.DayOfWeek;
            return workingHourConfig.WorkDays[(int)dayInWeek] == 1;
        }

        public static DateTime toLocaleTime(DateTime date, SDK.WorkingHours workingHourConfig) // date in utc, return local
        {
            return date.AddMinutes(-workingHourConfig.TimeZoneBias);
        }
    }
}
