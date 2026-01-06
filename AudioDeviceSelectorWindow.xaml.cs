using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace AutoStarter
{
    public partial class AudioDeviceSelectorWindow : Wpf.Ui.Controls.FluentWindow
    {
        private readonly List<DeviceInfo> _allPlaybackDevices;
        private readonly List<DeviceInfo> _allRecordingDevices;

        public DeviceInfo? SelectedDevice { get; private set; }

        public AudioDeviceSelectorWindow(IEnumerable<DeviceInfo> playbackDevices, IEnumerable<DeviceInfo> recordingDevices)
        {
            InitializeComponent();
            _allPlaybackDevices = playbackDevices.ToList();
            _allRecordingDevices = recordingDevices.ToList();
            UpdateDeviceLists();
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

        private void DeviceFilterCheckBoxChanged(object sender, RoutedEventArgs e)
        {
            UpdateDeviceLists();
        }

        private void UpdateDeviceLists()
        {
            bool showDisabled = ShowDisabledCheckBox.IsChecked == true;
            bool showUnplugged = ShowUnpluggedCheckBox.IsChecked == true;

            PlaybackDeviceListBox.ItemsSource = FilterDevices(_allPlaybackDevices, showDisabled, showUnplugged);
            RecordingDeviceListBox.ItemsSource = FilterDevices(_allRecordingDevices, showDisabled, showUnplugged);
        }

        private static IEnumerable<DeviceInfo> FilterDevices(IEnumerable<DeviceInfo> devices, bool showDisabled, bool showUnplugged)
        {
            return devices.Where(device =>
                device.State == NAudio.CoreAudioApi.DeviceState.Active ||
                (showDisabled && device.State == NAudio.CoreAudioApi.DeviceState.Disabled) ||
                (showUnplugged && (device.State == NAudio.CoreAudioApi.DeviceState.Unplugged || device.State == NAudio.CoreAudioApi.DeviceState.NotPresent)));
        }
    }
}
