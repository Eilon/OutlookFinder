using Newtonsoft.Json;
using System;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
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

            var userSettings = GetUserSettings();
            if (userSettings == null)
            {
                userSettings = CreateNewUserSettings();
                SaveUserSettings(userSettings);
            }
            OnRefreshSettings(userSettings);

            var thing = new OutlookFinderThing(userSettings);
            var results = thing.DoFind();

            _totalEmailsValueLabel.Text = results.TotalEmails.ToString(CultureInfo.CurrentCulture);
            _taggedEmailsValueLabel.Text = results.InterestingItems.Count.ToString(CultureInfo.CurrentCulture);

            var matchesPerTag = userSettings.SearchTerms
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

        private UserSettings CreateNewUserSettings()
        {
            using var userSettingsForm = new UserSettingsForm();
            //userSettingsForm.SetUserSettings(null);
            var result = userSettingsForm.ShowDialog(this);
            if (result == DialogResult.OK)
            {
                return userSettingsForm.GetUserSettings();
            }
            return null;
        }

        private void SaveUserSettings(UserSettings userSettings)
        {
            var userSettingsFileContents = JsonConvert.SerializeObject(userSettings);
            var userSettingsFilePath = GetUserSettingsFilePath();
            var userSettingsFolder = Path.GetDirectoryName(userSettingsFilePath);
            Directory.CreateDirectory(userSettingsFolder);
            File.WriteAllText(userSettingsFilePath, userSettingsFileContents);
        }

        private UserSettings GetUserSettings()
        {
            var userSettingsFilePath = GetUserSettingsFilePath();
            if (!File.Exists(userSettingsFilePath))
            {
                return null;
            }
            var userSettingsFileContents = File.ReadAllText(userSettingsFilePath);
            var userSettings = JsonConvert.DeserializeObject<UserSettings>(userSettingsFileContents);
            return userSettings;
        }

        private static string GetUserSettingsFilePath()
        {
            return Path.Combine(Application.UserAppDataPath, "OutlookFinder", "userSettings.json");
        }

        private void OnRefreshSettings(UserSettings userSettings)
        {
            _folderValueLabel.Text = string.Join(" / ", userSettings?.SearchFolderPath);
        }

        private void OutlookFinderAppForm_Load(object sender, EventArgs e)
        {
            OnRefreshSettings(GetUserSettings());
            _totalEmailsValueLabel.Text = "?";
            _taggedEmailsValueLabel.Text = "?";

            // Adjust the contents of the split container panels because for some reason setting the size in
            // the designer makes them too large at runtime.
            _tagResultsListView.Size = new Size(splitContainer1.Panel1.Width, splitContainer1.Panel1.Height) - new Size(10, 36);
            _logOutputTextBox.Size = new Size(splitContainer1.Panel2.Width, splitContainer1.Panel2.Height) - new Size(10, 36);
        }

        private void _settingsButton_Click(object sender, EventArgs e)
        {
            using var userSettingsForm = new UserSettingsForm();
            userSettingsForm.SetUserSettings(GetUserSettings());
            var result = userSettingsForm.ShowDialog(this);
            if (result == DialogResult.OK)
            {
                var newUserSettings = userSettingsForm.GetUserSettings();
                SaveUserSettings(newUserSettings);
                OnRefreshSettings(newUserSettings);
            }
        }
    }
}
