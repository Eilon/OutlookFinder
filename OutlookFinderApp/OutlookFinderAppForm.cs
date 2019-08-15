using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookFinderApp
{
    public partial class OutlookFinderAppForm : Form
    {
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

            var thing = new OutlookFinderThing(_folderToSearch);
            var results =  thing.DoFind();

            _totalEmailsValueLabel.Text = results.TotalEmails.ToString();
            _taggedEmailsValueLabel.Text = results.InterestingItems.Count.ToString();

            var allTags = results.InterestingItems
                .SelectMany(i => i.FoundSubstrings)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            var orderedTagCounts = allTags
                .Select(tag => (tag: tag, count: results.InterestingItems.Count(item => item.FoundSubstrings.Contains(tag, StringComparer.OrdinalIgnoreCase))))
                .OrderByDescending(tagCount => tagCount.count)
                .ToList();

            foreach (var (tag, count) in orderedTagCounts)
            {
                _tagResultsListView.Items.Add(new ListViewItem(new[] { tag, $"{count}" }));
            }
            _logOutputTextBox.Text = results.OutputLog.ToString();
        }

        private string _folderToSearch = "";

        private void OutlookFinderAppForm_Load(object sender, EventArgs e)
        {
            _folderValueLabel.Text = _folderToSearch;
            _totalEmailsValueLabel.Text = "?";
            _taggedEmailsValueLabel.Text = "?";

            // Adjust the contents of the split container panels because for some reason setting the size in
            // the designer makes them too large at runtime.
            _tagResultsListView.Size = new Size(splitContainer1.Panel1.Width, splitContainer1.Panel1.Height) - new Size(10, 36);
            _logOutputTextBox.Size = new Size(splitContainer1.Panel2.Width, splitContainer1.Panel2.Height) - new Size(10, 36);
        }
    }
}
