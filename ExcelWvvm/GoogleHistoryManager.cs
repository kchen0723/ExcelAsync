using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Net;

namespace ExcelWvvm
{
    public class GoogleHistoryManager
    {
        private const int MAX_GLITCH_CHECKS = 5;
        private static double yahooGlitchDate = new DateTime(2011, 1, 28).ToOADate();
        private static Dictionary<char, string> quoteHistoryParams = new Dictionary<char, string>() {
            {'d', "Date"},
            {'o', "Open"},
            {'h', "High"},
            {'l', "Low"},
            {'c', "Close"},
            {'v', "Volume"},
            {'a', "Adj Close"}
        };

        public static object[,] GoogleHistory(string secId, DateTime dblStartDate, DateTime dblEndDate)
        {
            return GoogleHistory(secId, dblStartDate.ToOADate(), dblEndDate.ToOADate(), "d", "dohlcv", false, false);
        }

        private static object[,] GoogleHistory(string secId, double dblStartDate, double dblEndDate, string period, string names, bool showHeaders, bool dateOrder)
        {
            DateTime startDate = (dblStartDate == 0) ? DateTime.Today.AddYears(-1) : DateTime.FromOADate(dblStartDate);
            DateTime endDate = (dblEndDate == 0) ? DateTime.Today : DateTime.FromOADate(dblEndDate);
            names = names.Replace('a', 'c');
            switch (ShortenDate(period))
            {
                case "w":
                    period = "weekly";
                    break;
                case "d":
                default:
                    period = "daily";
                    break;
            }
            string url = string.Format("http://www.google.com/finance/historical?q={0}&startdate={1}&enddate={2}&histperiod={3}&output=csv",
                secId, startDate.ToString("MMM+d,+yyyy"), endDate.ToString("MMM+d,+yyyy"), period);
            return QuoteHistoryParse(url, "d-MMM-yy", names, showHeaders, dateOrder, false);
        }

        private static object[,] QuoteHistoryParse(string url, string dateFormat, string names, bool showHeaders, bool dateOrder, bool checkGlitch)
        {
            // Used for determining whether to start on the second row of the CSV file when parsing
            int headerOffset = showHeaders ? 1 : 0;

            if (names != "")
            {
                object[,] csvResults = ImportCSV(url, 0, dateOrder, new string[] { dateFormat, "double" }, true);
                int counter = 0;

                // Fix the super-mega-weird Yahoo! glitch that randomly fails to return data after January 28, 2011
                // by requesting the same CSV again
                if (checkGlitch)
                {
                    while (counter++ < MAX_GLITCH_CHECKS && csvResults[headerOffset, 0].Equals(yahooGlitchDate))
                    {
                        csvResults = ImportCSV(url, 0, dateOrder, new string[] { dateFormat, "double" }, true);
                    }
                }

                // Fill out a list of headers so that we can easily find the text we're looking for and get the appropriate index
                Dictionary<string, int> headers = new Dictionary<string, int>();

                // Get height and width of CSV file
                int rowCount = csvResults.GetLength(0);
                int columnCount = csvResults.Length / rowCount;

                // Add all headers to a string array for storage
                for (int i = 0; i < columnCount; i++)
                {
                    headers.Add(csvResults[0, i].ToString(), i);
                }

                // Convert parameter names to an array of single characters
                char[] nameChars = names.ToLower().ToCharArray();
                object[,] results = new object[rowCount - (1 - headerOffset), nameChars.Length];
                int currentColumn = 0;
                foreach (char nameChar in nameChars)
                {
                    if (quoteHistoryParams.ContainsKey(nameChar) && headers.ContainsKey(quoteHistoryParams[nameChar]))
                    {
                        int matchColumn = headers[quoteHistoryParams[nameChar]];
                        for (int i = 1 - headerOffset; i < rowCount; i++)
                        {
                            results[i - (1 - headerOffset), currentColumn] = csvResults[i, matchColumn];
                        }
                    }
                    else
                    {
                        for (int i = 1 - headerOffset; i < rowCount; i++)
                        {
                            results[i - (1 - headerOffset), currentColumn] = 0;
                        }
                    }
                    currentColumn++;
                }

                return results;
            }

            return ImportCSV(url, 1 - headerOffset, dateOrder, new string[] { dateFormat, "double" }, showHeaders);
        }

        private static object[,] ImportCSV(string url, int startLine, bool reverse, object[] formats, bool hasHeaders)
        {
            CsvParseFormat formatter = new CsvParseFormat();
            WebRequest request;
            HttpWebResponse response;
            List<string[]> sorted = new List<string[]>();
            object[,] parsed;
            int counter = 0;

            foreach (object format in formats)
            {
                formatter.AddFormat(format.ToString());
            }

            request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = "GET";
            response = (HttpWebResponse)request.GetResponse();

            try
            {
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    string line;
                    string[] row;

                    while ((line = reader.ReadLine()) != null)
                    {
                        if (counter >= startLine)
                        {
                            row = line.Split(',');
                            sorted.Add(row);
                        }
                        counter++;
                    }
                }
            }
            catch (Exception e)
            {
                System.Windows.MessageBox.Show(e.Message);
            }

            if (reverse)
            {
                sorted.Reverse(hasHeaders ? 1 : 0, sorted.Count - (hasHeaders ? 1 : 0));
            }

            parsed = new object[sorted.Count, sorted[0].Length];

            for (int i = 0; i < sorted.Count; i++)
            {
                for (int j = 0; j < sorted[i].Length; j++)
                {
                    // Don't bother parsing headers
                    if (hasHeaders && i == 0)
                    {
                        parsed[i, j] = sorted[i][j].ToString();
                    }
                    else
                    {
                        try
                        {
                            parsed[i, j] = formatter.Parse(j, sorted[i][j]);
                        }
                        catch (Exception)
                        {
                            // parsed[i, j] = "";
                        }

                    }

                }
            }

            return parsed;
        }

        private class CsvParseFormat
        {
            private List<string> _formats = new List<string>();
            public void AddFormat(string format)
            {
                _formats.Add(format);
            }
            public object Parse(int key, string unparsed)
            {
                string format;

                if (key < _formats.Count)
                {
                    format = _formats[key];
                }
                else
                {
                    format = _formats.Last();
                }

                switch (format)
                {
                    case "string":
                        return unparsed;
                    case "double":
                        double tempVal;
                        double.TryParse(unparsed, out tempVal);
                        return tempVal;
                    default:
                        DateTime date = DateTime.ParseExact(unparsed, format, null);
                        return date.ToOADate();
                }
            }
        }

        private static string ShortenDate(string dateText)
        {
            switch (dateText.ToLower().Replace(" ", ""))
            {
                case "d":
                case "day":
                case "dly":
                case "daily":
                    return "d";
                case "ww":
                case "wwed":
                case "wkwed":
                case "weeklywednesday":
                case "weekly(wednesday)":
                    return "ww";
                case "wt":
                case "wthurs":
                case "wkthurs":
                case "weeklythursday":
                case "weekly(thursday)":
                    return "wt";
                case "wf":
                case "wfri":
                case "wkfri":
                case "weeklyfriday":
                case "weekly(friday)":
                    return "wf";
                case "w":
                case "wk":
                case "wkly":
                case "week":
                case "weekly":
                    return "w";
                case "bw":
                case "bwk":
                case "bweek":
                case "biweekly":
                case "bi-weekly":
                    return "bw";
                case "m":
                case "mth":
                case "month":
                case "monthly":
                    return "m";
                case "y":
                case "yr":
                case "year":
                case "yearly":
                case "a":
                case "ann":
                case "annual":
                    return "y";
                default:
                    return "";
            }
        }
    }
}
