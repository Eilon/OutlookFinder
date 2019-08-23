using System;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace OutlookFinderApp
{
    public partial class OutlookFinderAppForm : Form
    {
        private string[] _searchFolderPath = {
            "elipton@microsoft.com",
            "Inbox",
            "SellBuy",
        };

        private static readonly string[] SearchTerms = new[]
        {
            "infant",
            "baby",
            "toddler",
            "kid",
            "stroller",

            "toy",

            "bmw",
            "lego",
            "iphone",
            "apple",

            "shelf",
            "chair",
            "desk",
            "table",
            "dresser",
            "patio",
        };

        public OutlookFinderAppForm()
        {
            InitializeComponent();
        }

        private void OnRunNowButton_Click(object sender, EventArgs e)
        {
            _tagResultsListView.Items.Clear();
            _logOutputTextBox.Clear();
            _totalEmailsValueLabel.Text = "?";
            _taggedEmailsValueLabel.Text = "?";

            var thing = new OutlookFinderThing(_searchFolderPath, SearchTerms);
            var results = thing.DoFind();

            _totalEmailsValueLabel.Text = results.TotalEmails.ToString(CultureInfo.CurrentCulture);
            _taggedEmailsValueLabel.Text = results.InterestingItems.Count.ToString(CultureInfo.CurrentCulture);

            var matchesPerTag = SearchTerms
                .Select(searchTerm =>
                new
                {
                    searchTerm = searchTerm,
                    count = results.InterestingItems.Count(item => item.FoundSubstrings.Contains(searchTerm, StringComparer.OrdinalIgnoreCase))
                })
                .OrderByDescending(tagMatch => tagMatch.count)
                .ToList();


            foreach (var tagMatch in matchesPerTag)
            {
                _tagResultsListView.Items.Add(new ListViewItem(new[] { tagMatch.searchTerm, tagMatch.count.ToString(CultureInfo.CurrentCulture) }));
            }
            _logOutputTextBox.Text = results.OutputLog.ToString();
        }

        private void OutlookFinderAppForm_Load(object sender, EventArgs e)
        {
            _folderValueLabel.Text = string.Join(" / ", _searchFolderPath);
            _totalEmailsValueLabel.Text = "?";
            _taggedEmailsValueLabel.Text = "?";

            // Adjust the contents of the split container panels because for some reason setting the size in
            // the designer makes them too large at runtime.
            _tagResultsListView.Size = new Size(splitContainer1.Panel1.Width, splitContainer1.Panel1.Height) - new Size(10, 36);
            _logOutputTextBox.Size = new Size(splitContainer1.Panel2.Width, splitContainer1.Panel2.Height) - new Size(10, 36);
        }
    }
}
