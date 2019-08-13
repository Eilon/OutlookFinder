using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace OutlookFinderApp
{
    class OutlookFinderThing
    {
        private static readonly string[] SearchItems = new[]
        {
            "infant",
            "baby",
            "toddler",
            "kid",

            "toy",

            "bmw",
            "lego",
            "iphone",
            "apple",

            "shelf",
            "desk",
            "table",
            "stool",
            "dresser",
        };

        public OutlookFinderThing(string folderToSearch)
        {
            FolderToSearch = folderToSearch;
        }

        public string FolderToSearch { get; }

        public OutlookFindResults DoFind()
        {
            var results = new OutlookFindResults();

            Application myApp;
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {
                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                myApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application") as Application;
            }
            else
            {
                //if not, creating a new application instance
                myApp = new Application();
            }
            var mapiNameSpace = myApp.GetNamespace("MAPI");

            mapiNameSpace.Logon("", "", Missing.Value, Missing.Value);

            //var inboxFolder = mapiNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            //Console.WriteLine("Folders: {0}", inboxFolder.Folders.Count);

            // TODO: Use FolderToSearch
            var sellBuyFolder = mapiNameSpace.Folders["elipton@microsoft.com"].Folders["Inbox"].Folders["SellBuy"];

            results.TotalEmails = sellBuyFolder.Items.Count;

            results.OutputLog.AppendLine($"Total items in {sellBuyFolder.FullFolderPath} folder: {sellBuyFolder.Items.Count}");
            results.OutputLog.AppendLine("-----------------");
            var sellBuyEmails = sellBuyFolder.Items
                .OfType<MailItem>()
                .Select(m => new MailInfo(
                    m.SenderEmailAddress,
                    m.SenderName,
                    m.ReceivedTime,
                    m.Subject,
                    m.Body,
                    m.To,
                    m
                ))
                .ToList();

            var interestingItems = sellBuyEmails
                .Select(m => new InterestingMatch(m, FindInterestingMatches(m.Subject, m.Body, m.To)))
                .Where(match => match.FoundSubstrings.Any())
                .ToList();

            foreach (var interestingItem in interestingItems)
            {
                results.OutputLog.AppendLine($"{interestingItem.MailInfo.FromName} is selling '{interestingItem.MailInfo.Subject}' on {interestingItem.MailInfo.Received}");
                results.OutputLog.AppendLine($"\tKeywords: {string.Join(", ", interestingItem.FoundSubstrings)}");
                results.OutputLog.AppendLine($"\tPrices: {string.Join(", ", FindPrices(interestingItem.MailInfo.Subject).Select(p => $"${p}"))}");
                results.OutputLog.AppendLine($"\tPrices: {string.Join(", ", FindPrices(interestingItem.MailInfo.Body).Select(p => $"${p}"))}");

                var existingCategories =
                    interestingItem.MailInfo.MailItem.Categories?
                    .Split(';', ',')
                    .Select(c => c.Trim())
                    .Where(c => !string.IsNullOrWhiteSpace(c))
                    ?? Array.Empty<string>();
                interestingItem.MailInfo.MailItem.Categories = string.Join("; ", interestingItem.FoundSubstrings.Concat(existingCategories).Distinct());
                interestingItem.MailInfo.MailItem.Save();

                results.InterestingItems.Add(interestingItem);
            }

            return results;
        }

        private static string[] FindInterestingMatches(params string[] texts)
        {
            return texts
                .Where(text => text != null)
                .SelectMany(text => SearchItems.Where(searchItem => text.IndexOf(searchItem, StringComparison.OrdinalIgnoreCase) != -1))
                .ToArray();
        }

        private static decimal[] FindPrices(string text)
        {
            var priceRegex = new Regex(@"\$\d+");
            return priceRegex
                .Matches(text)
                .OfType<Match>()
                .SelectMany(match => match
                    .Captures
                    .OfType<Capture>()
                    .Select(c => decimal.TryParse(c.Value.Substring(1), out var result) ? result : (decimal?)null)
                    .Where(d => d.HasValue)
                    .Select(d => d.Value)
                )
                .ToArray();
        }
    }

    public class InterestingMatch
    {
        public InterestingMatch(MailInfo m, string[] foundSubstrings)
        {
            MailInfo = m;
            FoundSubstrings = foundSubstrings;
        }

        public string[] FoundSubstrings { get; }
        public MailInfo MailInfo { get; }
    }

    public class MailInfo
    {
        public string FromAddress { get; }
        public string FromName { get; }
        public DateTime Received { get; }
        public string Subject { get; }
        public string Body { get; }
        public string To { get; }
        public MailItem MailItem { get; }

        public MailInfo(string fromAddress, string fromName, DateTime received, string subject, string body, string to, MailItem mailItem)
        {
            FromAddress = fromAddress;
            FromName = fromName;
            Received = received;
            Subject = subject;
            Body = body;
            To = to;
            MailItem = mailItem;
        }

        public override bool Equals(object obj)
        {
            return obj is MailInfo other &&
                   FromAddress == other.FromAddress &&
                   FromName == other.FromName &&
                   Received == other.Received &&
                   Subject == other.Subject &&
                   Body == other.Body &&
                   To == other.To;
        }

        public override int GetHashCode()
        {
            var hashCode = -2016783565;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(FromAddress);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(FromName);
            hashCode = hashCode * -1521134295 + Received.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Subject);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Body);
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(To);
            return hashCode;
        }
    }
}
