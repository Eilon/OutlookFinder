using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OutlookFinder
{
    class Program
    {
        static void Main(string[] args)
        {
            Microsoft.Office.Interop.Outlook.Application myApp;
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {
                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                myApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
            }
            else
            {
                //if not, creating a new application instance
                myApp = new Microsoft.Office.Interop.Outlook.Application();
            }
            var mapiNameSpace = myApp.GetNamespace("MAPI");

            mapiNameSpace.Logon("", "", Missing.Value, Missing.Value);

            //var inboxFolder = mapiNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            //Console.WriteLine("Folders: {0}", inboxFolder.Folders.Count);

            var sellBuyFolder = mapiNameSpace.Folders["elipton@microsoft.com"].Folders["Inbox"].Folders["SellBuy"];

            Console.WriteLine($"Total items in {sellBuyFolder.FullFolderPath} folder: {sellBuyFolder.Items.Count}");
            Console.WriteLine("-----------------");
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
                //.Take(100) // TODO: Remove this
                .ToList();

            var interestingItems = sellBuyEmails
                .Select(m => new InterestingMatch(m, FindInterestingMatches(m.Subject, m.Body, m.To)))
                .Where(match => match.FoundSubstrings.Any())
                .ToList();

            var outputBuilder = new StringBuilder();

            RenderHeader(outputBuilder, SearchItems);

            foreach (var interestingItem in interestingItems)
            {
                RenderItem(outputBuilder, interestingItem, SearchItems);
                Console.WriteLine($"{interestingItem.MailInfo.FromName} is selling '{interestingItem.MailInfo.Subject}' on {interestingItem.MailInfo.Received}");
                Console.WriteLine($"\tKeywords: {string.Join(", ", interestingItem.FoundSubstrings)}");
                Console.WriteLine($"\tPrices: {string.Join(", ", FindPrices(interestingItem.MailInfo.Subject).Select(p => $"${p}"))}");
                Console.WriteLine($"\tPrices: {string.Join(", ", FindPrices(interestingItem.MailInfo.Body).Select(p => $"${p}"))}");
                //Console.WriteLine("Accounts: {0}", mailItem.Body);

                var existingCategories =
                    interestingItem.MailInfo.MailItem.Categories
                    .Split(';', ',')
                    .Select(c => c.Trim())
                    .Where(c => !string.IsNullOrWhiteSpace(c));
                interestingItem.MailInfo.MailItem.Categories = string.Join("; ", interestingItem.FoundSubstrings.Concat(existingCategories).Distinct());
                interestingItem.MailInfo.MailItem.Save();
            }

            RenderFooter(outputBuilder, SearchItems);

            File.WriteAllText("out.html", outputBuilder.ToString());

            Console.ReadLine();
        }

        private static void RenderHeader(StringBuilder outputBuilder, string[] searchItems)
        {
            outputBuilder.AppendLine($@"<table id=""table_id"" class=""display"">
    <thead>
        <tr>
            <th>From</th>
            <th>To</th>
            <th>Subject</th>
            <th>Date</th>
            <th>Subject Prices</th>
            <th>Body Prices</th>
{string.Join("", searchItems.Select(item => $"            <th>{item}</th>{Environment.NewLine}"))}
        </tr>
    </thead>
    <tbody>
");
        }

        private static void RenderItem(StringBuilder outputBuilder, InterestingMatch interestingItem, string[] searchItems)
        {
            outputBuilder.AppendLine($@"
        <tr>
            <td>{interestingItem.MailInfo.FromName}</td>
            <td>{interestingItem.MailInfo.To}</td>
            <td>{interestingItem.MailInfo.Subject}</td>
            <td>{interestingItem.MailInfo.Received}</td>
            <td>{string.Join(", ", FindPrices(interestingItem.MailInfo.Subject).Select(p => $"${p}"))}</td>
            <td>{string.Join(", ", FindPrices(interestingItem.MailInfo.Body).Select(p => $"${p}"))}</td>
{string.Join("", searchItems.Select(item => $"            <td>{interestingItem.FoundSubstrings.Contains(item)}</td>{Environment.NewLine}"))}
        </tr>");
        }

        private static void RenderFooter(StringBuilder outputBuilder, string[] searchItems)
        {
            outputBuilder.AppendLine($@"
    </tbody>
</table>");
        }

        private static readonly string[] SearchItems = new[]
        {
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
        };

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

    public static class TypeInformation
    {
        public static string GetTypeName(object comObject)
        {
            var dispatch = comObject as IDispatch;

            if (dispatch == null)
            {
                return null;
            }

            var pTypeInfo = dispatch.GetTypeInfo(0, 1033);

            string pBstrName;
            string pBstrDocString;
            int pdwHelpContext;
            string pBstrHelpFile;
            pTypeInfo.GetDocumentation(
                -1,
                out pBstrName,
                out pBstrDocString,
                out pdwHelpContext,
                out pBstrHelpFile);

            string str = pBstrName;
            if (str[0] == 95)
            {
                // remove leading '_'
                str = str.Substring(1);
            }

            return str;
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("00020400-0000-0000-C000-000000000046")]
        private interface IDispatch
        {
            int GetTypeInfoCount();

            [return: MarshalAs(UnmanagedType.Interface)]
            ITypeInfo GetTypeInfo(
                [In, MarshalAs(UnmanagedType.U4)] int iTInfo,
                [In, MarshalAs(UnmanagedType.U4)] int lcid);

            void GetIDsOfNames(
                [In] ref Guid riid,
                [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames,
                [In, MarshalAs(UnmanagedType.U4)] int cNames,
                [In, MarshalAs(UnmanagedType.U4)] int lcid,
                [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        }
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
