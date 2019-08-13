using System.Collections.Generic;
using System.Text;

namespace OutlookFinderApp
{
    internal class OutlookFindResults
    {
        public int TotalEmails { get; set; }

        public StringBuilder OutputLog { get; } = new StringBuilder();
        public List<InterestingMatch> InterestingItems { get; } = new List<InterestingMatch>();
    }
}
