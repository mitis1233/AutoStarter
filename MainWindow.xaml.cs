using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using Microsoft.Win32;
using NAudio.CoreAudioApi;

namespace AutoStarter;

public partial class MainWindow : Window
{
    private static readonly JsonSerializerOptions _jsonSerializerOptions = new()
    {
        Converters = { new JsonStringEnumConverter() },
        WriteIndented = true
    };

    public ObservableCollection<ActionItem> ActionItems { get; set; }


    public MainWindow()
    {
        InitializeComponent();
        ActionItems = [];

        DataContext = this;
    }

    private void AddApp_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "可執行檔 (*.exe)|*.exe|所有檔案 (*.*)|*.*"
        };

        if (openFileDialog.ShowDialog() == true)
        {
            ActionItems.Add(new ActionItem
            {
                Type = ActionType.LaunchApplication,
                FilePath = openFileDialog.FileName,
                Arguments = ""
            });
        }
    }

    private void AddDelay_Click(object sender, RoutedEventArgs e)
    {
        ActionItems.Add(new ActionItem { Type = ActionType.Delay, DelaySeconds = 5 });
    }

    private async void AddAudio_Click(object sender, RoutedEventArgs e)
    {
        var enumerator = new MMDeviceEnumerator();

        var (playbackDevices, recordingDevices) = await Task.Run(() =>
        {
            var playback = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                                     .Select(d => new DeviceInfo { ID = d.ID, FriendlyName = d.FriendlyName })
                                     .ToList();
            var recording = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
                                      .Select(d => new DeviceInfo { ID = d.ID, FriendlyName = d.FriendlyName })
                                      .ToList();
            return (playback, recording);
        });

        if (playbackDevices.Count == 0 && recordingDevices.Count == 0)
        {
            MessageBox.Show("找不到任何已啟用的音訊裝置。", "提示", MessageBoxButton.OK, MessageBoxImage.None);
            return;
        }

        var selectorWindow = new AudioDeviceSelectorWindow(playbackDevices, recordingDevices);
        if (selectorWindow.ShowDialog() == true && selectorWindow.SelectedDevice != null)
        {
            var selectedDevice = selectorWindow.SelectedDevice;
            ActionItems.Add(new ActionItem
            {
                Type = ActionType.SetAudioDevice,
                AudioDeviceId = selectedDevice.ID,
                AudioDeviceName = selectedDevice.FriendlyName
            });
        }
    }

    private async void AddDisableAudio_Click(object sender, RoutedEventArgs e)
    {
        var enumerator = new MMDeviceEnumerator();

        var (playbackDevices, recordingDevices) = await Task.Run(() =>
        {
            var playback = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                                     .Select(d => new DeviceInfo { ID = d.ID, FriendlyName = d.FriendlyName })
                                     .ToList();
            var recording = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
                                      .Select(d => new DeviceInfo { ID = d.ID, FriendlyName = d.FriendlyName })
                                      .ToList();
            return (playback, recording);
        });

        if (playbackDevices.Count == 0 && recordingDevices.Count == 0)
        {
            MessageBox.Show("找不到任何已啟用的音訊裝置。", "提示", MessageBoxButton.OK, MessageBoxImage.None);
            return;
        }

        var selectorWindow = new AudioDeviceSelectorWindow(playbackDevices, recordingDevices);
        if (selectorWindow.ShowDialog() == true && selectorWindow.SelectedDevice != null)
        {
            var selectedDevice = selectorWindow.SelectedDevice;
            ActionItems.Add(new ActionItem
            {
                Type = ActionType.DisableAudioDevice,
                AudioDeviceId = selectedDevice.ID,
                AudioDeviceName = selectedDevice.FriendlyName
            });
        }
    }

    private void MoveUp_Click(object sender, RoutedEventArgs e)
    {
        var selectedIndex = ActionsDataGrid.SelectedIndex;
        if (selectedIndex > 0)
        {
            ActionItems.Move(selectedIndex, selectedIndex - 1);
        }
    }

    private void MoveDown_Click(object sender, RoutedEventArgs e)
    {
        var selectedIndex = ActionsDataGrid.SelectedIndex;
        if (selectedIndex != -1 && selectedIndex < ActionItems.Count - 1)
        {
            ActionItems.Move(selectedIndex, selectedIndex + 1);
        }
    }

    private void Remove_Click(object sender, RoutedEventArgs e)
    {
        // Make a copy of the selected items to avoid issues with modifying the collection while iterating.
        var selectedItems = ActionsDataGrid.SelectedItems.Cast<ActionItem>().ToList();
        foreach (var item in selectedItems)
        {
            ActionItems.Remove(item);
        }
    }



    private void Delete_Click(object sender, RoutedEventArgs e)
    {
        if (ActionsDataGrid.SelectedItem is not ActionItem selectedItem) return;

        if (ActionsDataGrid.ItemsSource is ObservableCollection<ActionItem> items)
        {
            items.Remove(selectedItem);
        }
    }

    private async void Edit_Click(object sender, RoutedEventArgs e)
    {
        if (ActionsDataGrid.SelectedItem is not ActionItem selectedItem) return;

        switch (selectedItem.Type)
        {
            case ActionType.SetAudioDevice:
            case ActionType.DisableAudioDevice:
                var enumerator = new MMDeviceEnumerator();
                var (playbackDevices, recordingDevices) = await Task.Run(() =>
                {
                    var playback = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active)
                                             .Select(d => new DeviceInfo { ID = d.ID, FriendlyName = d.FriendlyName })
                                             .ToList();
                    var recording = enumerator.EnumerateAudioEndPoints(DataFlow.Capture, DeviceState.Active)
                                              .Select(d => new DeviceInfo { ID = d.ID, FriendlyName = d.FriendlyName })
                                              .ToList();
                    return (playback, recording);
                });

                var selectorWindow = new AudioDeviceSelectorWindow(playbackDevices, recordingDevices) { Owner = this };
                if (selectorWindow.ShowDialog() == true && selectorWindow.SelectedDevice != null)
                {
                    selectedItem.AudioDeviceId = selectorWindow.SelectedDevice.ID;
                    selectedItem.AudioDeviceName = selectorWindow.SelectedDevice.FriendlyName;
                }
                break;

            default:
                var editWindow = new EditActionWindow(selectedItem)
                {
                    Owner = this
                };
                editWindow.ShowDialog();
                break;
        }
    }

    private void ImportProfile_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "AutoStarter 設定檔 (*.autostart)|*.autostart|所有檔案 (*.*)|*.*",
            Title = "匯入設定檔"
        };

        if (openFileDialog.ShowDialog() == true)
        {
            try
            {
                var jsonString = File.ReadAllText(openFileDialog.FileName);
                var importedItems = JsonSerializer.Deserialize<List<ActionItem>>(jsonString, _jsonSerializerOptions);

                if (importedItems != null)
                {
                    foreach (var item in importedItems)
                    {
                        ActionItems.Add(item);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"匯入設定檔時發生錯誤：\n{ex.Message}", "匯入失敗", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private async void SaveProfile_Click(object sender, RoutedEventArgs e)
    {
        var saveFileDialog = new SaveFileDialog
        {
            Filter = "啟動設定檔 (*.autostart)|*.autostart",
            DefaultExt = "autostart",
            FileName = "MyStartup.autostart"
        };

        if (saveFileDialog.ShowDialog() == true)
        {
            var jsonString = JsonSerializer.Serialize(ActionItems, _jsonSerializerOptions);
            await File.WriteAllTextAsync(saveFileDialog.FileName, jsonString);
            MessageBox.Show("設定檔已儲存！", "成功", MessageBoxButton.OK, MessageBoxImage.None);
        }
    }

    private void RegisterFileAssociation_Click(object sender, RoutedEventArgs e)
    {
        try
        {
                            string? exePath = Environment.ProcessPath;
                if (string.IsNullOrEmpty(exePath))
                {
                    MessageBox.Show("無法取得主程式路徑。", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            RegistryKey key = Registry.ClassesRoot.CreateSubKey(".autostart");
            key.SetValue("", "AutoStarter.Profile");
            key.Close();

            key = Registry.ClassesRoot.CreateSubKey("AutoStarter.Profile\\shell\\open\\command");
            key.SetValue("", $"\"{exePath}\" \"%1\"");
            key.Close();

            MessageBox.Show("檔案關聯已成功註冊！", "成功", MessageBoxButton.OK, MessageBoxImage.None);
        }
        catch (System.Exception ex)
        {
            MessageBox.Show($"註冊檔案關聯失敗：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.None);
        }
    }

    private void UnregisterFileAssociation_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            Registry.ClassesRoot.DeleteSubKeyTree(".autostart", false);
            Registry.ClassesRoot.DeleteSubKeyTree("AutoStarter.Profile", false);
            MessageBox.Show("檔案關聯已成功移除！", "成功", MessageBoxButton.OK, MessageBoxImage.None);
        }
        catch (System.Exception ex)
        {
            MessageBox.Show($"移除檔案關聯失敗：{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.None);
        }
    }

    private void ActionsDataGrid_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {

    }
}
