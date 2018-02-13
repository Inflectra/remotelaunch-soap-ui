using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Inflectra.RemoteLaunch.Engines.soapUI
{
    /// <summary>
    /// Interaction logic for soapUISettings.xaml
    /// </summary>
    public partial class SoapUISettings : UserControl
    {
        public SoapUISettings()
        {
            InitializeComponent();
            this.LoadSettings();
        }

        /// <summary>
        /// Loads the saved settings
        /// </summary>
        private void LoadSettings()
        {
            //Load the various properties
            this.txtLocation.Text = Properties.Settings.Default.Location;
            this.chkSoapUiPro.IsChecked = Properties.Settings.Default.ProLicense;
            this.chkLoudUi.IsChecked = Properties.Settings.Default.LoadTest;
            this.chkTraceLogging.IsChecked = Properties.Settings.Default.TraceLogging;
        }

        /// <summary>
        /// Saves the specified settings.
        /// </summary>
        public void SaveSettings()
        {
            //Get the various properties
            Properties.Settings.Default.Location = this.txtLocation.Text.Trim();
            Properties.Settings.Default.ProLicense = this.chkSoapUiPro.IsChecked.Value;
            Properties.Settings.Default.LoadTest = this.chkLoudUi.IsChecked.Value;
            Properties.Settings.Default.TraceLogging = this.chkTraceLogging.IsChecked.Value;

            //Save the properties and reload
            Properties.Settings.Default.Save();
            this.LoadSettings();
        }

        /// <summary>
        /// Displays the file selection dialog when clicked
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            //Create the new file open dialog
            OpenFileDialog fileOpen = new OpenFileDialog();
            fileOpen.MultiSelect = false;
            fileOpen.DefaultExt = "bat";
            fileOpen.CheckFileExists = true;
            fileOpen.CheckPathExists = true;
            fileOpen.Filter = "*.bat";
            Nullable<bool> result = fileOpen.ShowDialog(Window.GetWindow(this));
            if (result.HasValue && result.Value)
            {
                //Set the text box to the selected file
                this.txtLocation.Text = System.IO.Path.GetDirectoryName(fileOpen.FileName);
            }
        }
    }
}
