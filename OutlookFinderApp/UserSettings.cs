using System.Collections.Generic;

namespace OutlookFinderApp
{
    public class UserSettings
    {
#pragma warning disable CA2227 // Collection properties should be read only
        public IList<string> SearchFolderPath { get; set; }
        public IList<string> SearchTerms { get; set; }
#pragma warning restore CA2227 // Collection properties should be read only
    }
}
