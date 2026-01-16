using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Threading;
using AutoStarter.CoreAudio;
using NAudio.CoreAudioApi;

namespace AutoStarter;

public partial class MainWindow : Wpf.Ui.Controls.FluentWindow
{
    public class PowerPlan
    {
        public string Name { get; set; } = "";
        public Guid Guid { get; set; } 
    }

    public ObservableCollection<ActionItem> ActionItems { get; set; }
    private string? _lastImportedProfileDirectory;
    private string? _lastExportDirectory;
    private bool _suppressAutoScroll;

    private Task<Wpf.Ui.Controls.MessageBoxResult> ShowDialogAsync(string title, string content, string primaryText = "確定", string? closeText = null)
    {
        return ThemedDialogService.ShowAsync(this, title, content, primaryText, closeText);
    }

    private async Task<DeviceInfo?> PromptForAudioDeviceAsync(string? preselectedDeviceId = null)
    {
        var (playbackDevices, recordingDevices) = AudioDeviceService.GetDeviceLists(DeviceState.All);

        if (playbackDevices.Count == 0 && recordingDevices.Count == 0)
        {
            await ShowDialogAsync("提示", "找不到任何音訊裝置。");
            return null;
        }

        var selectorWindow = new AudioDeviceSelectorWindow(playbackDevices, recordingDevices, preselectedDeviceId)
        {
            Owner = this
        };

        return selectorWindow.ShowDialog() == true ? selectorWindow.SelectedDevice : null;
    }

    public MainWindow()
    {
        InitializeComponent();
        ActionItems = [];
        DataContext = this;
    }

    protected override void OnContentRendered(EventArgs e)
    {
        base.OnContentRendered(e);
        ActionItems.CollectionChanged += ActionItems_CollectionChanged;
    }

    private void ActionItems_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (_suppressAutoScroll)
        {
            return;
        }

        if (e.Action != NotifyCollectionChangedAction.Add || e.NewItems == null || e.NewItems.Count == 0)
        {
            return;
        }

        var newestItem = e.NewItems[^1];

        Dispatcher.BeginInvoke(() =>
        {
            if (ActionsDataGrid == null || !IsLoaded)
            {
                return;
            }

            ActionsDataGrid.UpdateLayout();
            ActionsDataGrid.ScrollIntoView(newestItem);
            ActionsDataGrid.SelectedItem = newestItem;
        }, DispatcherPriority.Background);
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
        var selectedDevice = await PromptForAudioDeviceAsync();
        if (selectedDevice == null)
        {
            return;
        }

        ActionItems.Add(new ActionItem
        {
            Type = ActionType.SetAudioDevice,
            AudioDeviceId = selectedDevice.ID,
            AudioDeviceInstanceId = selectedDevice.InstanceId,
            AudioDeviceName = selectedDevice.FriendlyName
        });
    }

    private async void AddPowerPlan_Click(object sender, RoutedEventArgs e)
    {
        var powerPlans = await LoadPowerPlans();
        if (powerPlans.Count == 0)
        {
            await ShowDialogAsync("錯誤", "找不到任何電源計畫。");
            return;
        }

        var selectorWindow = new PowerPlanSelectorWindow(powerPlans) { Owner = this };
        if (selectorWindow.ShowDialog() == true && selectorWindow.SelectedPlan != null)
        {
            ActionItems.Add(new ActionItem
            {
                Type = ActionType.SetPowerPlan,
                PowerPlanId = selectorWindow.SelectedPlan.Guid,
                PowerPlanName = selectorWindow.SelectedPlan.Name
            });
        }
    }

    private async void AddDisableAudio_Click(object sender, RoutedEventArgs e)
    {
        var selectedDevice = await PromptForAudioDeviceAsync();
        if (selectedDevice == null)
        {
            return;
        }

        ActionItems.Add(new ActionItem
        {
            Type = ActionType.DisableAudioDevice,
            AudioDeviceId = selectedDevice.ID,
            AudioDeviceInstanceId = selectedDevice.InstanceId,
            AudioDeviceName = selectedDevice.FriendlyName
        });
    }

    private async void AddEnableAudio_Click(object sender, RoutedEventArgs e)
    {
        var selectedDevice = await PromptForAudioDeviceAsync();
        if (selectedDevice == null)
        {
            return;
        }

        ActionItems.Add(new ActionItem
        {
            Type = ActionType.EnableAudioDevice,
            AudioDeviceId = selectedDevice.ID,
            AudioDeviceInstanceId = selectedDevice.InstanceId,
            AudioDeviceName = selectedDevice.FriendlyName
        });
    }

    private void AddAudioVolume_Click(object sender, RoutedEventArgs e)
    {
        var playbackVolume = TryGetDefaultEndpointVolume(DataFlow.Render);
        var recordingVolume = TryGetDefaultEndpointVolume(DataFlow.Capture);

        var volumeWindow = new AudioVolumeWindow(
            adjustPlaybackVolume: playbackVolume.HasValue,
            playbackVolumePercent: playbackVolume ?? 50,
            adjustRecordingVolume: false,
            recordingVolumePercent: recordingVolume ?? 50)
        {
            Owner = this
        };

        if (volumeWindow.ShowDialog() == true)
        {
            var action = new ActionItem
            {
                Type = ActionType.SetAudioVolume,
                AdjustPlaybackVolume = volumeWindow.AdjustPlaybackVolume,
                PlaybackVolumePercent = volumeWindow.AdjustPlaybackVolume ? volumeWindow.PlaybackVolumePercent : null,
                AdjustRecordingVolume = volumeWindow.AdjustRecordingVolume,
                RecordingVolumePercent = volumeWindow.AdjustRecordingVolume ? volumeWindow.RecordingVolumePercent : null
            };

            action.AudioVolumePercent = action.AdjustPlaybackVolume
                ? action.PlaybackVolumePercent
                : action.AdjustRecordingVolume
                    ? action.RecordingVolumePercent
                    : null;

            ActionItems.Add(action);
        }
    }

    private static int? TryGetDefaultEndpointVolume(DataFlow flow)
    {
        try
        {
            using var enumerator = new MMDeviceEnumerator();
            var device = enumerator.GetDefaultAudioEndpoint(flow, Role.Multimedia);
            if (device?.AudioEndpointVolume == null)
            {
                return null;
            }

            var scalar = device.AudioEndpointVolume.MasterVolumeLevelScalar;
            return Math.Clamp((int)Math.Round(scalar * 100), 0, 100);
        }
        catch
        {
            return null;
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
            case ActionType.EnableAudioDevice:
            case ActionType.SetPowerPlan:
                if (selectedItem.Type == ActionType.SetPowerPlan)
                {
                    var powerPlans = await LoadPowerPlans();
                    var powerPlanSelector = new PowerPlanSelectorWindow(powerPlans) { Owner = this };
                    if (powerPlanSelector.ShowDialog() == true && powerPlanSelector.SelectedPlan != null)
                    {
                        selectedItem.PowerPlanId = powerPlanSelector.SelectedPlan.Guid;
                        selectedItem.PowerPlanName = powerPlanSelector.SelectedPlan.Name;
                    }
                    return; // Early exit for power plan
                }

                var (playbackDevices, recordingDevices) = AudioDeviceService.GetDeviceLists(DeviceState.All);

                var selectorWindow = new AudioDeviceSelectorWindow(playbackDevices, recordingDevices) { Owner = this };
                if (selectorWindow.ShowDialog() == true && selectorWindow.SelectedDevice != null)
                {
                    selectedItem.AudioDeviceId = selectorWindow.SelectedDevice.ID;
                    selectedItem.AudioDeviceInstanceId = selectorWindow.SelectedDevice.InstanceId;
                    selectedItem.AudioDeviceName = selectorWindow.SelectedDevice.FriendlyName;
                }
                break;

            case ActionType.SetAudioVolume:
                var initialPlaybackPercent = selectedItem.PlaybackVolumePercent
                    ?? selectedItem.AudioVolumePercent ?? 50;
                var initialRecordingPercent = selectedItem.RecordingVolumePercent
                    ?? selectedItem.AudioVolumePercent ?? 50;

                var volumeWindow = new AudioVolumeWindow(
                    adjustPlaybackVolume: selectedItem.AdjustPlaybackVolume,
                    playbackVolumePercent: initialPlaybackPercent,
                    adjustRecordingVolume: selectedItem.AdjustRecordingVolume,
                    recordingVolumePercent: initialRecordingPercent)
                {
                    Owner = this
                };

                if (volumeWindow.ShowDialog() == true)
                {
                    selectedItem.AdjustPlaybackVolume = volumeWindow.AdjustPlaybackVolume;
                    selectedItem.PlaybackVolumePercent = volumeWindow.AdjustPlaybackVolume
                        ? volumeWindow.PlaybackVolumePercent
                        : null;

                    selectedItem.AdjustRecordingVolume = volumeWindow.AdjustRecordingVolume;
                    selectedItem.RecordingVolumePercent = volumeWindow.AdjustRecordingVolume
                        ? volumeWindow.RecordingVolumePercent
                        : null;

                    selectedItem.AudioVolumePercent = selectedItem.AdjustPlaybackVolume
                        ? selectedItem.PlaybackVolumePercent
                        : selectedItem.AdjustRecordingVolume
                            ? selectedItem.RecordingVolumePercent
                            : null;
                }
                return;

            default:
                var editWindow = new EditActionWindow(selectedItem)
                {
                    Owner = this
                };
                editWindow.ShowDialog();
                break;
        }
    }

    private async void ImportProfile_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "AutoStarter 設定檔 (*.autostart)|*.autostart|所有檔案 (*.*)|*.*",
            Title = "匯入設定檔"
        };

        if (openFileDialog.ShowDialog() == true)
        {
            await ImportProfilesAsync(new[] { openFileDialog.FileName });
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

        saveFileDialog.InitialDirectory = _lastImportedProfileDirectory ?? _lastExportDirectory;

        if (saveFileDialog.ShowDialog() == true)
        {
            await AutostartProfileService.SaveAsync(saveFileDialog.FileName, ActionItems);
            _lastExportDirectory = Path.GetDirectoryName(saveFileDialog.FileName);
        }
    }

    private async void RegisterFileAssociation_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            string? exePath = Environment.ProcessPath;
            if (string.IsNullOrEmpty(exePath))
            {
                await ShowDialogAsync("錯誤", "無法取得主程式路徑。");
                return;
            }

            using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(".autostart"))
            {
                if (key != null)
                {
                    key.SetValue("", "AutoStarter.Profile");
                }
            }

            using (RegistryKey key = Registry.ClassesRoot.CreateSubKey("AutoStarter.Profile\\shell\\open\\command"))
            {
                if (key != null)
                {
                    key.SetValue("", $"\"{exePath}\" \"%1\"");
                }
            }

            await ShowDialogAsync("成功", "檔案關聯已成功註冊！");
        }
        catch (System.Exception ex)
        {
            await ShowDialogAsync("錯誤", $"註冊檔案關聯失敗：{ex.Message}");
        }
    }

    private async void UnregisterFileAssociation_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            // DeleteSubKeyTree 不需要 using，因為它不返回 RegistryKey
            Registry.ClassesRoot.DeleteSubKeyTree(".autostart", false);
            Registry.ClassesRoot.DeleteSubKeyTree("AutoStarter.Profile", false);
            await ShowDialogAsync("成功", "檔案關聯已成功移除！");
        }
        catch (System.Exception ex)
        {
            await ShowDialogAsync("錯誤", $"移除檔案關聯失敗：{ex.Message}");
        }
    }
        private async Task<List<PowerPlan>> LoadPowerPlans()
    {
        var powerPlans = new List<PowerPlan>();
        var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "powercfg",
                Arguments = "/list",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
            }
        };

        process.Start();
        string output = await process.StandardOutput.ReadToEndAsync();
        await process.WaitForExitAsync();

        var regex = new Regex(@"GUID: (.*?)  \((.*?)\)");
        var matches = regex.Matches(output);

        foreach (Match match in matches)
        {
            if (Guid.TryParse(match.Groups[1].Value.Trim(), out Guid guid))
            {
                powerPlans.Add(new PowerPlan { Guid = guid, Name = match.Groups[2].Value.Trim() });
            }
        }

        return powerPlans;
    }

    private void ActionsDataGrid_DragOver(object sender, DragEventArgs e)

    {
        // 檢查是否有檔案被拖動
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            // 只有當至少有一個.autostart檔案時才允許拖放
            if (files != null && files.Any(f => f.EndsWith(".autostart", StringComparison.OrdinalIgnoreCase)))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
                return;
            }
        }
        e.Effects = DragDropEffects.None;
        e.Handled = true;
    }

    private async void ActionsDataGrid_Drop(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null)
            {
                var autostartFiles = files
                    .Where(file => file.EndsWith(".autostart", StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                if (autostartFiles.Length > 0)
                {
                    await ImportProfilesAsync(autostartFiles);
                }
            }
        }
        e.Handled = true;
    }

    private async Task ImportProfilesAsync(IEnumerable<string> filePaths)
    {
        var importedItems = new List<ActionItem>();

        foreach (var filePath in filePaths)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    await ShowDialogAsync("錯誤", $"檔案不存在：{filePath}");
                    continue;
                }

                var items = await AutostartProfileService.LoadAsync(filePath);
                RememberImportedDirectory(filePath);

                if (items.Count > 0)
                {
                    importedItems.AddRange(items);
                }
            }
            catch (Exception ex)
            {
                await ShowDialogAsync("匯入失敗", $"匯入設定檔時發生錯誤：\n{ex.Message}");
            }
        }

        AddActionItems(importedItems);
    }

    private void AddActionItems(IReadOnlyList<ActionItem> items)
    {
        if (items == null || items.Count == 0)
        {
            return;
        }

        _suppressAutoScroll = true;
        try
        {
            foreach (var item in items)
            {
                ActionItems.Add(item);
            }
        }
        finally
        {
            _suppressAutoScroll = false;
        }

        var newestItem = items[^1];
        Dispatcher.BeginInvoke(() =>
        {
            if (ActionsDataGrid == null || !IsLoaded)
            {
                return;
            }

            ActionsDataGrid.UpdateLayout();
            ActionsDataGrid.ScrollIntoView(newestItem);
            ActionsDataGrid.SelectedItem = newestItem;
        }, DispatcherPriority.Background);
    }

    private void RemoveAll_Click(object sender, RoutedEventArgs e)
    {
        if (ActionItems.Count > 0)
        {
            ActionItems.Clear();
        }
    }

    private void RememberImportedDirectory(string filePath)
    {
        try
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                _lastImportedProfileDirectory = directory;
            }
        }
        catch
        {
            // ignore invalid path formats
        }
    }

}
