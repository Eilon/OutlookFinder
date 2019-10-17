using System;
using System.Windows.Forms;

namespace OutlookFinderApp
{
    public partial class UserSettingsForm : Form
    {
        public UserSettingsForm()
        {
            InitializeComponent();
        }

        public void SetUserSettings(UserSettings userSettings)
        {
            if (userSettings is null)
            {
                throw new ArgumentNullException(nameof(userSettings));
            }
            _searchFolderTextBox.Text = string.Join(Environment.NewLine, userSettings.SearchFolderPath);
            _searchTermsTextBox.Text = string.Join(Environment.NewLine, userSettings.SearchTerms);
        }

        public UserSettings GetUserSettings()
        {
            return new UserSettings
            {
                SearchFolderPath = _searchFolderTextBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries),
                SearchTerms = _searchTermsTextBox.Text.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries),
            };
        }

        private void _okButton_Click(object sender, EventArgs e)
        {
            // TODO: Consider running validation code here?
        }

        private void _cancelButton_Click(object sender, EventArgs e)
        {

        }
    }
}
