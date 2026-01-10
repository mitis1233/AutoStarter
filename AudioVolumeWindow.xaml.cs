using System;
using System.Windows;

namespace AutoStarter
{
    public partial class AudioVolumeWindow : Wpf.Ui.Controls.FluentWindow
    {
        public bool AdjustPlaybackVolume { get; private set; }
        public bool AdjustRecordingVolume { get; private set; }
        public int PlaybackVolumePercent { get; private set; }
        public int RecordingVolumePercent { get; private set; }

        private bool _isUpdatingPlaybackValue;
        private bool _isUpdatingRecordingValue;
        private bool _isLoaded;

        public AudioVolumeWindow(
            bool adjustPlaybackVolume = true,
            int playbackVolumePercent = 50,
            bool adjustRecordingVolume = false,
            int recordingVolumePercent = 50)
        {
            InitializeComponent();

            AdjustPlaybackVolume = adjustPlaybackVolume;
            PlaybackVolumePercent = Math.Clamp(playbackVolumePercent, 0, 100);
            AdjustRecordingVolume = adjustRecordingVolume;
            RecordingVolumePercent = Math.Clamp(recordingVolumePercent, 0, 100);

            Loaded += OnLoaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            _isLoaded = true;

            PlaybackCheckBox.IsChecked = AdjustPlaybackVolume;
            RecordingCheckBox.IsChecked = AdjustRecordingVolume;

            PlaybackSlider.Value = PlaybackVolumePercent;
            PlaybackValueBox.Text = PlaybackVolumePercent.ToString();
            RecordingSlider.Value = RecordingVolumePercent;
            RecordingValueBox.Text = RecordingVolumePercent.ToString();

            UpdatePlaybackUi();
            UpdateRecordingUi();
        }

        private void PlaybackSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isLoaded || _isUpdatingPlaybackValue)
                return;

            _isUpdatingPlaybackValue = true;
            PlaybackVolumePercent = (int)Math.Round(e.NewValue);
            PlaybackValueBox.Text = PlaybackVolumePercent.ToString();
            _isUpdatingPlaybackValue = false;
        }

        private void RecordingSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (!_isLoaded || _isUpdatingRecordingValue)
                return;

            _isUpdatingRecordingValue = true;
            RecordingVolumePercent = (int)Math.Round(e.NewValue);
            RecordingValueBox.Text = RecordingVolumePercent.ToString();
            _isUpdatingRecordingValue = false;
        }

        private void PlaybackCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (!_isLoaded)
                return;

            AdjustPlaybackVolume = PlaybackCheckBox.IsChecked == true;
            UpdatePlaybackUi();
        }

        private void RecordingCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (!_isLoaded)
                return;

            AdjustRecordingVolume = RecordingCheckBox.IsChecked == true;
            UpdateRecordingUi();
        }

        private void UpdatePlaybackUi()
        {
            if (PlaybackSliderHost != null)
            {
                PlaybackSliderHost.Visibility = AdjustPlaybackVolume ? Visibility.Visible : Visibility.Collapsed;
            }

            PlaybackValueBox.IsEnabled = AdjustPlaybackVolume;
            PlaybackValueBox.Opacity = AdjustPlaybackVolume ? 1 : 0.4;
        }

        private void UpdateRecordingUi()
        {
            if (RecordingSliderHost != null)
            {
                RecordingSliderHost.Visibility = AdjustRecordingVolume ? Visibility.Visible : Visibility.Collapsed;
            }

            RecordingValueBox.IsEnabled = AdjustRecordingVolume;
            RecordingValueBox.Opacity = AdjustRecordingVolume ? 1 : 0.4;
        }

        private void PlaybackValueBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (!_isLoaded || _isUpdatingPlaybackValue)
                return;

            if (!int.TryParse(PlaybackValueBox.Text, out var value))
                return;

            value = Math.Clamp(value, 0, 100);

            _isUpdatingPlaybackValue = true;
            if (PlaybackSlider.Value != value)
            {
                PlaybackSlider.Value = value;
            }
            PlaybackValueBox.Text = value.ToString();
            PlaybackVolumePercent = value;
            _isUpdatingPlaybackValue = false;
        }

        private void RecordingValueBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (!_isLoaded || _isUpdatingRecordingValue)
                return;

            if (!int.TryParse(RecordingValueBox.Text, out var value))
                return;

            value = Math.Clamp(value, 0, 100);

            _isUpdatingRecordingValue = true;
            if (RecordingSlider.Value != value)
            {
                RecordingSlider.Value = value;
            }
            RecordingValueBox.Text = value.ToString();
            RecordingVolumePercent = value;
            _isUpdatingRecordingValue = false;
        }

        private async void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            if (!AdjustPlaybackVolume && !AdjustRecordingVolume)
            {
                await ThemedDialogService.ShowAsync(this, "提示", "請至少選擇一個要調整的裝置。");
                return;
            }

            if (AdjustPlaybackVolume)
            {
                PlaybackVolumePercent = (int)Math.Round(PlaybackSlider.Value);
            }

            if (AdjustRecordingVolume)
            {
                RecordingVolumePercent = (int)Math.Round(RecordingSlider.Value);
            }

            DialogResult = true;
        }
    }
}
