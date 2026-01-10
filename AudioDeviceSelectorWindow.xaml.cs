using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace AutoStarter
{
    public partial class AudioDeviceSelectorWindow : Wpf.Ui.Controls.FluentWindow
    {
        private readonly List<DeviceInfo> _allPlaybackDevices;
        private readonly List<DeviceInfo> _allRecordingDevices;
        private readonly string? _preselectedDeviceId;
        private bool _preselectionApplied;

        public DeviceInfo? SelectedDevice { get; private set; }

        public AudioDeviceSelectorWindow(
            IEnumerable<DeviceInfo> playbackDevices,
            IEnumerable<DeviceInfo> recordingDevices,
            string? preselectedDeviceId = null)
        {
            InitializeComponent();
            _allPlaybackDevices = playbackDevices.ToList();
            _allRecordingDevices = recordingDevices.ToList();
            _preselectedDeviceId = preselectedDeviceId;
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

            ApplyPreselectionIfNeeded();
        }

        private static IEnumerable<DeviceInfo> FilterDevices(IEnumerable<DeviceInfo> devices, bool showDisabled, bool showUnplugged)
        {
            return devices.Where(device =>
                device.State == NAudio.CoreAudioApi.DeviceState.Active ||
                (showDisabled && device.State == NAudio.CoreAudioApi.DeviceState.Disabled) ||
                (showUnplugged && (device.State == NAudio.CoreAudioApi.DeviceState.Unplugged || device.State == NAudio.CoreAudioApi.DeviceState.NotPresent)));
        }

        private void ApplyPreselectionIfNeeded()
        {
            if (_preselectionApplied || string.IsNullOrWhiteSpace(_preselectedDeviceId))
            {
                return;
            }

            bool playbackSelected = TrySelectDevice(PlaybackDeviceListBox);
            bool recordingSelected = TrySelectDevice(RecordingDeviceListBox);

            _preselectionApplied = playbackSelected || recordingSelected;
        }

        private bool TrySelectDevice(System.Windows.Controls.ListBox listBox)
        {
            if (listBox.ItemsSource is not IEnumerable<DeviceInfo> devices)
            {
                return false;
            }

            var match = devices.FirstOrDefault(device =>
                device.ID != null &&
                string.Equals(device.ID, _preselectedDeviceId, StringComparison.OrdinalIgnoreCase));

            if (match == null)
            {
                return false;
            }

            listBox.SelectedItem = match;
            listBox.ScrollIntoView(match);
            return true;
        }
    }
}
