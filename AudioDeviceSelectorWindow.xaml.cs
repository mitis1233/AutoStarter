using System.Windows;

namespace AutoStarter
{
    public partial class AudioDeviceSelectorWindow : Wpf.Ui.Controls.FluentWindow
    {
        public DeviceInfo? SelectedDevice { get; private set; }

        public AudioDeviceSelectorWindow(IEnumerable<DeviceInfo> playbackDevices, IEnumerable<DeviceInfo> recordingDevices)
        {
            InitializeComponent();
            PlaybackDeviceListBox.ItemsSource = playbackDevices;
            RecordingDeviceListBox.ItemsSource = recordingDevices;
        }

        private void DeviceListBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            var selectedListBox = sender as System.Windows.Controls.ListBox;
            if (selectedListBox?.SelectedItem != null)
            {
                if (selectedListBox == PlaybackDeviceListBox)
                {
                    RecordingDeviceListBox.UnselectAll();
                }
                else
                {
                    PlaybackDeviceListBox.UnselectAll();
                }
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (PlaybackDeviceListBox.SelectedItem != null)
            {
                SelectedDevice = (DeviceInfo)PlaybackDeviceListBox.SelectedItem;
                DialogResult = true;
            }
            else if (RecordingDeviceListBox.SelectedItem != null)
            {
                SelectedDevice = (DeviceInfo)RecordingDeviceListBox.SelectedItem;
                DialogResult = true;
            }
            else
            {
                MessageBox.Show("請選擇一個裝置。", "提示", MessageBoxButton.OK, MessageBoxImage.None);
            }
        }
    }
}
