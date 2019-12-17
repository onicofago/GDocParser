using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Docs.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;

namespace DocsQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/docs.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { DocsService.Scope.DocumentsReadonly };
        static string ApplicationName = "Google Docs API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Docs API service.
            var service = new DocsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String documentId = "10iaemvyMan4ZfjUj1VqI1ubPCsh-krpPI92Pm3z5zjs";
            DocumentsResource.GetRequest request = service.Documents.Get(documentId);

            // Parse the document's body 
            Document doc = request.Execute();
            Console.WriteLine($"Title: {doc.Title}");
            var paragraphs = 0;
            var days = 0;
            var wfh = 0;
            var tagDict = new Dictionary<string, int>();
            foreach (var element in doc.Body.Content)
            {
                if (element.Paragraph != null)
                {
                    paragraphs += 1;
                    var sb = new StringBuilder();
                    foreach (var parel in element.Paragraph.Elements)
                    {
                        sb.Append(parel.TextRun?.Content);
                    }
                    var text = sb.ToString();

                    // Date "title" parsing
                    if (IsADateHeading(text))
                    {
                        Console.WriteLine($"Date: {GetDateFromTitle(text.Trim()).ToString("dd/MMM/yyyy")}");
                        days += 1;
                        if (text.ToLower().Contains("[wfh]")) wfh += 1;
                        SearchForPossibleDayTags(text, tagDict);
                    }

                    // Pluralsight example
                    if ((text != null) && (text.Contains("Pluralsight:"))) Console.WriteLine(text);
                }
            }
            Console.WriteLine($"\nParagraph count: {paragraphs}");
            Console.WriteLine($"Work days: {days}");
            Console.WriteLine($"Remoting days: {wfh}");
            Console.WriteLine("\nDay Tags:");
            foreach (var item in tagDict.OrderBy(i => i.Key))
            {
                Console.WriteLine($"[{item.Key}]: {item.Value}");
            }
            Console.ReadLine();
        }

        private static DateTime GetDateFromTitle(string title)
        {
            title = title.Substring(0, 8);
            var dateString = new String(title.Where(Char.IsDigit).ToArray());
            return DateTime.ParseExact(dateString, "yyyyMMdd", CultureInfo.InvariantCulture);
        }

        private static void SearchForPossibleDayTags(string text, Dictionary<string, int> dict)
        {
            var tags = new List<string>();
            var bInsideTag = false;
            var currentTag = "";
            foreach (var c in text.ToCharArray())
            {
                switch (c)
                {
                    case '[':
                        bInsideTag = true;
                        break;
                    case ']':
                        bInsideTag = false;
                        currentTag = currentTag.Trim();
                        if (!tags.Contains(currentTag)) tags.Add(currentTag);
                        currentTag = "";
                        break;
                    default:
                        if (bInsideTag) currentTag += c.ToString();
                        break;
                }
            }
            foreach (var tag in tags)
            {
                if (!dict.ContainsKey(tag))
                    dict.Add(tag, 1);
                else
                    dict[tag] += 1;
            }
        }

        private static bool IsADateHeading(string text)
        {
            if (text == null) return false;
            return Regex.IsMatch(text.Trim(), @"^([2-9]\d{3}((0[1-9]|1[012])(0[1-9]|1\d|2[0-8])|(0[13456789]|1[012])(29|30)|(0[13578]|1[02])31)|(([2-9]\d)(0[48]|[2468][048]|[13579][26])|(([2468][048]|[3579][26])00))0229)\b");
        }
    }
}
