using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
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
                //Log($"找到 .autostart 檔案：{filePath}");
                try
                {
                    string json = File.ReadAllText(filePath);
                    JsonSerializerOptions jsonSerializerOptions = new()
                    {
                        Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() }
                    };
                    var options = jsonSerializerOptions;
                    var actions = JsonSerializer.Deserialize<List<ActionItem>>(json, options);
                    if (actions != null)
                    {
                        foreach (var action in actions)
                        {
                            switch (action.Type)
                            {
                                case ActionType.LaunchApplication:
                                    //Log($"執行操作：啟動應用程式 - {action.FilePath}");
                                    if (!string.IsNullOrEmpty(action.FilePath) && File.Exists(action.FilePath))
                                    {
                                        var startInfo = new ProcessStartInfo(action.FilePath)
                                        {
                                            Arguments = action.Arguments,
                                            UseShellExecute = true
                                        };

                                        // First, try the standard way.
                                        if (action.MinimizeWindow)
                                        {
                                            startInfo.WindowStyle = ProcessWindowStyle.Minimized;
                                        }

                                        var process = Process.Start(startInfo);

                                        // If minimization is requested, wait for the window and force it if necessary.
                                        if (action.MinimizeWindow && process != null)
                                        {
                                            try
                                            {
                                                // Wait until the window handle is available, with a timeout.
                                                var stopwatch = Stopwatch.StartNew();
                                                const int timeoutMs = 15000;
                                                while (process.MainWindowHandle == IntPtr.Zero && stopwatch.ElapsedMilliseconds < timeoutMs)
                                                {
                                                    await Task.Delay(250);
                                                    process.Refresh();
                                                }

                                                if (process.MainWindowHandle != IntPtr.Zero)
                                                {
                                                    // If the window is not already minimized, force it.
                                                    if (!IsIconic(process.MainWindowHandle))
                                                    {
                                                        //Log($"應用程式 {action.FilePath} 的視窗在啟動時未最小化。正在強制最小化。");
                                                        ShowWindow(process.MainWindowHandle, SW_MINIMIZE);
                                                    }
                                                }
                                                else
                                                {
                                                    //Log($"在 {timeoutMs} 毫秒內找不到應用程式 {action.FilePath} 的主視窗控制代碼。");
                                                }
                                            }
                                            catch (Exception )//ex)
                                            {
                                                //Log($"嘗試最小化 {action.FilePath} 時發生錯誤：{ex.Message}");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        //Log($"找不到應用程式路徑或路徑無效：{action.FilePath}");
                                    }
                                    break;
                                case ActionType.Delay:
                                    //Log($"執行操作：延遲 {action.DelaySeconds} 秒");
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
                catch (Exception )//ex)
                {
                    //Log($"處理 .autostart 檔案時發生錯誤：{ex.Message}");
                }

                //Log("已完成處理 .autostart 檔案。正在關閉程式。");
                Shutdown();
            }
            else
            {
                var mainWindow = new MainWindow();
                mainWindow.Show();
            }
        }

        private static void SetDefaultAudioDevice(string deviceId)
        {
            //Log("執行操作：設定音訊裝置");
            //Log($"正在嘗試設定音訊裝置。ID：{deviceId}");
            try
            {
                var client = new PolicyConfigClient();
                //Log("正在為「主控台」設定預設端點。");
                client.SetDefaultEndpoint(deviceId, ERole.eConsole);
                //Log("正在為「多媒體」設定預設端點。");
                client.SetDefaultEndpoint(deviceId, ERole.eMultimedia);
                //Log("正在為「通訊」設定預設端點。");
                client.SetDefaultEndpoint(deviceId, ERole.eCommunications);
                //Log("成功設定音訊裝置。");
            }
            catch (Exception )//ex)
            {
                //Log($"設定音訊裝置失敗：{ex.Message}");
            }
        }

        private static void DisableAudioDevice(string deviceId)
        {
            //Log("執行操作：停用音訊裝置");
            //Log($"正在嘗試停用音訊裝置。ID：{deviceId}");
            try
            {
                var client = new PolicyConfigClient();
                //Log("正在停用音訊裝置。");
                client.DisableEndpoint(deviceId);
                //Log("成功停用音訊裝置。");
            }
            catch (Exception )//ex)
            {
                //Log($"停用音訊裝置失敗：{ex.Message}");
            }
        }

        private static void EnableAudioDevice(string deviceId)
        {
            //Log("執行操作：啟用音訊裝置");
            //Log($"正在嘗試啟用音訊裝置。ID：{deviceId}");
            try
            {
                var client = new PolicyConfigClient();
                //Log("正在啟用音訊端點。");
                client.EnableEndpoint(deviceId);
                //Log("成功啟用音訊裝置。");
            }
            catch (Exception )//ex)
            {
                //Log($"啟用音訊裝置失敗：{ex.Message}");
            }
        }

        // Win32 API functions
        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static partial bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static partial bool IsIconic(IntPtr hWnd);

        private const int SW_MINIMIZE = 6;

        //private static void Log(string message)
        //{
        //    try
        //    {
        //        string logFilePath = Path.Combine(AppContext.BaseDirectory, "autostarter.log");
        //        File.AppendAllText(logFilePath, $"{DateTime.Now:yyyy/M/d HH:mm:ss}: {message}{Environment.NewLine}");
        //    }
        //    catch
        //    {
        //        // Ignore logging errors
        //    }
        //}
    }
}
