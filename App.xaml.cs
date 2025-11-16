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

                                        // If minimization is requested, handle it asynchronously in the background
                                        if (action.MinimizeWindow && process != null)
                                        {
                                            // Fire and forget: minimize in background without blocking the main flow
                                            _ = MinimizeWindowAsync(process);
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

        private static async Task MinimizeWindowAsync(Process process)
        {
            if (process == null || process.HasExited)
                return;

            try
            {
                const int totalWaitMs = 5000;    // 總等待時間 5 秒
                const int checkIntervalMs = 100;  // 檢查間隔 100ms
                int elapsedMs = 0;
                
                // 等待視窗句柄可用
                while (process.MainWindowHandle == IntPtr.Zero && elapsedMs < totalWaitMs)
                {
                    await Task.Delay(checkIntervalMs);
                    elapsedMs += checkIntervalMs;
                    
                    try 
                    {
                        if (process.HasExited)
                            return;
                            
                        process.Refresh();
                    }
                    catch (InvalidOperationException)
                    {
                        // 進程已結束或無法存取
                        return;
                    }
                }

                // 如果還是沒有視窗句柄，放棄嘗試
                if (process.MainWindowHandle == IntPtr.Zero)
                    return;

                // 短暫等待確保視窗完全初始化
                await Task.Delay(300);

                // 最多重試 3 次，使用指數退避策略
                for (int attempt = 0; attempt < 3; attempt++)
                {
                    try
                    {
                        if (process.HasExited)
                            return;

                        process.Refresh();
                        
                        // 檢查視窗是否已經最小化
                        if (IsIconic(process.MainWindowHandle) || process.MainWindowHandle == IntPtr.Zero)
                            break;

                        // 嘗試最小化
                        if (!ShowWindow(process.MainWindowHandle, SW_MINIMIZE))
                        {
                            // 如果失敗，記錄錯誤碼（如果需要可以取消註解）
                            // int error = Marshal.GetLastWin32Error();
                            // Log($"Minimize failed with error code: {error}");
                        }
                        
                        // 指數退避：150ms, 300ms
                        if (attempt < 2)
                            await Task.Delay(150 * (int)Math.Pow(2, attempt));
                    }
                    catch (Exception) when (IsProcessExitedOrInaccessible(process))
                    {
                        // 進程已結束或無法存取，直接返回
                        return;
                    }
                    catch (Exception ex) when (ex is ObjectDisposedException || 
                                            ex is InvalidOperationException ||
                                            ex is System.ComponentModel.Win32Exception)
                    {
                        // 忽略這些已知的異常類型，並在最後一次嘗試時重新拋出
                        if (attempt == 2)
                            throw;
                    }
                }
            }
            catch (Exception ex)
            {
                // 記錄未捕獲的異常（如果需要可以取消註解）
                // Log($"Unexpected error in MinimizeWindowAsync: {ex}");
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
        
        // 檢查進程是否已結束或無法存取
        private static bool IsProcessExitedOrInaccessible(Process process)
        {
            if (process == null)
                return true;
                
            try
            {
                return process.HasExited;
            }
            catch
            {
                return true;
            }
        }

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
