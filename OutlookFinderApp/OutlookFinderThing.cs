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
        public OutlookFinderThing(UserSettings userSettings)
        {
            SearchFolderPath = userSettings.SearchFolderPath;
            SearchTerms = userSettings.SearchTerms;
        }

        public IList<string> SearchFolderPath { get; }
        public IList<string> SearchTerms { get; }

        public OutlookFindResults DoFind()
        {
            var results = new OutlookFindResults();

            Application myApp;
            if (Process.GetProcessesByName("OUTLOOK").Any())
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

            var searchFolder = GetSearchFolder(mapiNameSpace, SearchFolderPath);

            results.TotalEmails = searchFolder.Items.Count;

            results.OutputLog.AppendLine($"Total items in {searchFolder.FullFolderPath} folder: {searchFolder.Items.Count}");
            results.OutputLog.AppendLine("-----------------");
            var sellBuyEmails = searchFolder.Items
                .OfType<MailItem>()
                .Where(m => m.MessageClass == "IPM.Note") // IPM.Note means an email. See https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/bb176446(v=office.12)
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
                .Select(m => new InterestingMatch(m, FindInterestingMatches(SearchTerms, m.Subject, m.Body, m.To)))
                .Where(match => match.FoundSubstrings.Any())
                .ToList();

            foreach (var interestingItem in interestingItems)
            {
                results.OutputLog.AppendLine($"{interestingItem.MailInfo.FromName} is selling '{interestingItem.MailInfo.Subject}' on {interestingItem.MailInfo.Received}");
                results.OutputLog.AppendLine($"\tKeywords: {string.Join(", ", interestingItem.FoundSubstrings)}");
                var distinctSortedPrices =
                    FindPrices(interestingItem.MailInfo.Subject)
                    .Concat(FindPrices(interestingItem.MailInfo.Body))
                    .Distinct()
                    .OrderByDescending(d => d);

                var pricesString = distinctSortedPrices.Any() ? string.Join(", ", distinctSortedPrices.Select(p => $"${p}")) : "<none>";
                results.OutputLog.AppendLine($"\tPrices: {pricesString}");

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

        private static MAPIFolder GetSearchFolder(NameSpace rootNamespace, IList<string> searchFolderPath)
        {
            if (searchFolderPath.Count == 0)
            {
                throw new ArgumentException("Search folder path must have at least 1 segment.", nameof(searchFolderPath));
            }
            var firstSubfolder = rootNamespace.Folders[searchFolderPath[0]];
            var currentSubfolder = firstSubfolder;

            for (int i = 1; i < searchFolderPath.Count; i++)
            {
                currentSubfolder = currentSubfolder.Folders[searchFolderPath[i]];
            }
            return currentSubfolder;
        }

        private static string[] FindInterestingMatches(IList<string> searchTerms, params string[] texts)
        {
            return texts
                .Where(text => text != null)
                .SelectMany(text => searchTerms.Where(searchItem => text.IndexOf(searchItem, StringComparison.OrdinalIgnoreCase) != -1))
                .ToArray();
        }

        private static decimal[] FindPrices(string text)
        {
            if (text is null)
            {
                return Array.Empty<decimal>();
            }
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

#pragma warning disable CA1819 // Properties should not return arrays
        public string[] FoundSubstrings { get; }
#pragma warning restore CA1819 // Properties should not return arrays
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
