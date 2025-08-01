using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using AutoStarter.CoreAudio;
using System.Text.Json;

namespace AutoStarter
{
    public partial class App : Application
    {
        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            if (e.Args.Length > 0 && e.Args[0].EndsWith(".autostart"))
            {
                string filePath = e.Args[0];
                Log($"Found .autostart file: {filePath}");
                try
                {
                    string json = File.ReadAllText(filePath);
                    var options = new JsonSerializerOptions
                    {
                        Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() }
                    };
                    var actions = JsonSerializer.Deserialize<List<ActionItem>>(json, options);
                    if (actions != null)
                    {
                        foreach (var action in actions)
                        {
                            switch (action.Type)
                            {
                                case ActionType.LaunchApplication:
                                    Log($"Executing action: LaunchApp - {action.FilePath}");
                                    if (!string.IsNullOrEmpty(action.FilePath) && File.Exists(action.FilePath))
                                    {
                                        Process.Start(new ProcessStartInfo(action.FilePath) { Arguments = action.Arguments, UseShellExecute = true });
                                    }
                                    else
                                    {
                                        Log($"App path not found or invalid: {action.FilePath}");
                                    }
                                    break;
                                case ActionType.Delay:
                                    Log($"Executing action: Delay for {action.DelaySeconds} seconds");
                                    await Task.Delay(TimeSpan.FromSeconds(action.DelaySeconds));
                                    break;
                                case ActionType.SetAudioDevice:
                                    if (action.AudioDeviceId != null)
                                    {
                                        Dispatcher.Invoke(() =>
                                        {
                                            EnableAudioDevice(action.AudioDeviceId);
                                            SetDefaultAudioDevice(action.AudioDeviceId);
                                        });
                                    }
                                    break;
                                case ActionType.DisableAudioDevice:
                                    if (action.AudioDeviceId != null)
                                    {
                                        Dispatcher.Invoke(() => DisableAudioDevice(action.AudioDeviceId));
                                    }
                                    break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log($"Error processing .autostart file: {ex.Message}");
                }

                Log("Finished processing .autostart file. Shutting down.");
                Shutdown();
            }
            else
            {
                var mainWindow = new MainWindow();
                mainWindow.Show();
            }
        }

        private void SetDefaultAudioDevice(string deviceId)
        {
            Log("Executing action: SetAudioDevice");
            Log($"Attempting to set audio device. ID: {deviceId}");
            try
            {
                var client = new PolicyConfigClient();
                Log("Setting default endpoint for Console.");
                client.SetDefaultEndpoint(deviceId, ERole.eConsole);
                Log("Setting default endpoint for Multimedia.");
                client.SetDefaultEndpoint(deviceId, ERole.eMultimedia);
                Log("Setting default endpoint for Communications.");
                client.SetDefaultEndpoint(deviceId, ERole.eCommunications);
                Log("Successfully set audio device.");
            }
            catch (Exception ex)
            {
                Log($"Failed to set audio device: {ex.Message}");
            }
        }

        private void DisableAudioDevice(string deviceId)
        {
            Log("Executing action: DisableAudioDevice");
            Log($"Attempting to disable audio device. ID: {deviceId}");
            try
            {
                var client = new PolicyConfigClient();
                Log("Disabling audio device.");
                client.DisableEndpoint(deviceId);
                Log("Successfully disabled audio device.");
            }
            catch (Exception ex)
            {
                Log($"Failed to disable audio device: {ex.Message}");
            }
        }

        private void EnableAudioDevice(string deviceId)
        {
            Log("Executing action: EnableAudioDevice");
            Log($"Attempting to enable audio device. ID: {deviceId}");
            try
            {
                var client = new PolicyConfigClient();
                Log("Enabling audio endpoint.");
                client.EnableEndpoint(deviceId);
                Log("Successfully enabled audio device.");
            }
            catch (Exception ex)
            {
                Log($"Failed to enable audio device: {ex.Message}");
            }
        }

        private void Log(string message)
        {
            try
            {
                string logFilePath = Path.Combine(AppContext.BaseDirectory, "autostarter.log");
                File.AppendAllText(logFilePath, $"{DateTime.Now:yyyy/M/d HH:mm:ss}: {message}{Environment.NewLine}");
            }
            catch
            {
                // Ignore logging errors
            }
        }
    }
}
