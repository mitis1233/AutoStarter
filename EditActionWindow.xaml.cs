using System.Windows;
using System.Windows.Controls;

namespace AutoStarter
{
    public partial class EditActionWindow : Wpf.Ui.Controls.FluentWindow
    {
        private readonly ActionItem _actionItem;
        private TextBox? _argumentsTextBox;
        private TextBox? _delayTextBox;
        private TextBox? _filePathTextBox;

        public EditActionWindow(ActionItem item)
        {
            InitializeComponent();
            _actionItem = item;
            DataContext = _actionItem;

            SetupUI();
        }

        private void SetupUI()
        {
            EditorStackPanel.Children.Clear();

            switch (_actionItem.Type)
            {
                case ActionType.LaunchApplication:
                    // File Path
                    EditorStackPanel.Children.Add(new Label { Content = "檔案路徑:" });
                    var grid = new Grid();
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                    _filePathTextBox = new TextBox { Text = _actionItem.FilePath, Margin = new Thickness(0, 0, 5, 0) };
                    Grid.SetColumn(_filePathTextBox, 0);
                    grid.Children.Add(_filePathTextBox);

                    var browseButton = new Wpf.Ui.Controls.Button { Content = "瀏覽...", Padding = new Thickness(5,0,5,0) };
                    browseButton.Style = (Style)FindResource("FrostedGlassButtonStyle");
                    browseButton.Click += BrowseButton_Click;
                    Grid.SetColumn(browseButton, 1);
                    grid.Children.Add(browseButton);
                    EditorStackPanel.Children.Add(grid);

                    // Arguments
                    EditorStackPanel.Children.Add(new Label { Content = "啟動參數:" });
                    _argumentsTextBox = new TextBox { Text = _actionItem.Arguments };
                    EditorStackPanel.Children.Add(_argumentsTextBox);
                    break;

                case ActionType.Delay:
                    EditorStackPanel.Children.Add(new Label { Content = "延遲 (秒):" });
                    _delayTextBox = new TextBox { Text = _actionItem.DelaySeconds.ToString(), Width = 250 };
                    EditorStackPanel.Children.Add(_delayTextBox);
                    break;
                case ActionType.SetAudioVolume:
                    EditorStackPanel.Children.Add(new Label { Content = "音量百分比 (0-100):" });
                    _delayTextBox = new TextBox { Text = (_actionItem.AudioVolumePercent ?? 50).ToString(), Width = 250 };
                    EditorStackPanel.Children.Add(_delayTextBox);
                    break;
                
                default:
                    EditorStackPanel.Children.Add(new TextBlock { Text = "此項目類型沒有可編輯的屬性。", Margin = new Thickness(5) });
                    // Disable OK button if there's nothing to edit
                    if (FindName("OkButton") is Button okButton) { okButton.IsEnabled = false; }
                    break;
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "可執行檔 (*.exe)|*.exe|所有檔案 (*.*)|*.*",
                FileName = _filePathTextBox?.Text
            };

            if (openFileDialog.ShowDialog() == true)
            {
                if (_filePathTextBox != null) _filePathTextBox.Text = openFileDialog.FileName;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            switch (_actionItem.Type)
            {
                case ActionType.LaunchApplication:
                    if (_filePathTextBox != null) _actionItem.FilePath = _filePathTextBox.Text;
                    if (_argumentsTextBox != null) _actionItem.Arguments = _argumentsTextBox.Text;
                    break;

                case ActionType.Delay:
                    if (_delayTextBox != null && int.TryParse(_delayTextBox.Text, out int delay))
                    {
                        _actionItem.DelaySeconds = delay;
                    }
                    else
                    {
                        MessageBox.Show("請輸入有效的數字。", "輸入錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return; // Keep window open
                    }
                    break;
                case ActionType.SetAudioVolume:
                    if (_delayTextBox != null && int.TryParse(_delayTextBox.Text, out int volume) && volume is >= 0 and <= 100)
                    {
                        _actionItem.AudioVolumePercent = volume;
                    }
                    else
                    {
                        MessageBox.Show("請輸入 0 到 100 之間的整數。", "輸入錯誤", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    break;
            }

            DialogResult = true;
        }
    }
}
