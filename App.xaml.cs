using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using AutoStarter.CoreAudio;
using System.Text.Json;
using System.Management;
using System.Text;
using System.Linq;

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
                        var minimizeTasks = new List<Task>();
                        
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
                                        if (action.MinimizeWindow || action.ForceMinimizeWindow)
                                        {
                                            startInfo.WindowStyle = ProcessWindowStyle.Minimized;
                                        }

                                        var process = Process.Start(startInfo);

                                        // If minimization is requested, handle it asynchronously in the background
                                        if (process != null)
                                        {
                                            if (action.ForceMinimizeWindow)
                                            {
                                                // 強制最小化：使用監控模式
                                                minimizeTasks.Add(MinimizeWindowWithProcessMonitoringAsync(process));
                                            }
                                            else if (action.MinimizeWindow)
                                            {
                                                // 普通最小化
                                                minimizeTasks.Add(MinimizeWindowAsync(process));
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
                                case ActionType.SetPowerPlan:
                                    if (action.PowerPlanId != Guid.Empty)
                                    {
                                        var guid = action.PowerPlanId;
                                        PowerSetActiveScheme(IntPtr.Zero, ref guid);
                                    }
                                    break;
                            }
                        }
                        
                        // 等待所有最小化任務完成
                        if (minimizeTasks.Count > 0)
                        {
                            await Task.WhenAll(minimizeTasks);
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
            catch (Exception)
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

        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static partial bool IsWindow(IntPtr hWnd);

        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static partial bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [LibraryImport("user32.dll")]
        private static partial IntPtr GetWindow(IntPtr hWnd, int uCmd);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [LibraryImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static partial bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [LibraryImport("user32.dll", EntryPoint = "GetWindowLongPtrW", SetLastError = true)]
        private static partial IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

        private const int GWL_STYLE = -16;
        private const long WS_MINIMIZEBOX = 0x00020000L;

        private const int SW_MINIMIZE = 6;
        private const int GW_OWNER = 4;
        private const uint WM_SYSCOMMAND = 0x0112;
        private const int SC_MINIMIZE = 0xF020;

        [DllImport("powrprof.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint PowerSetActiveScheme(IntPtr UserPowerKey, ref Guid ActivePolicyGuid);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [Flags]
        private enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0,
            SMTO_BLOCK = 0x1,
            SMTO_ABORTIFHUNG = 0x2
        }

        private const uint WM_NULL = 0x0000;
        
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

        private static async Task MinimizeWindowWithProcessMonitoringAsync(Process launcherProcess)
        {
            if (launcherProcess == null)
                return;

            try
            {
                var mainWindowHandle = await WaitForMainWindowHandleAsync(launcherProcess, 15000, 250);
                
                // 記錄初始窗口
                var initialWindowHandles = new HashSet<IntPtr>(EnumerateAllVisibleWindows());
                if (mainWindowHandle != IntPtr.Zero)
                {
                    initialWindowHandles.Add(mainWindowHandle);
                }

                // 監控新創建的窗口
                IntPtr targetWindowHandle = IntPtr.Zero;
                DateTime lastNewWindowTime = DateTime.MinValue;
                bool foundNewWindow = false;  // 標記是否發現新窗口

                for (int i = 0; i < 20; i++)  // 最多等待 15 秒
                {
                    await Task.Delay(1000);  // 每秒檢查一次

                    try
                    {
                        var currentWindows = EnumerateAllVisibleWindows();
                        var newWindows = currentWindows.Where(h => !initialWindowHandles.Contains(h)).ToList();

                        if (newWindows.Any())
                        {
                            foreach (var windowHandle in newWindows)
                            {
                                try
                                {
                                    long style = GetWindowLongPtr(windowHandle, GWL_STYLE).ToInt64();
                                    bool hasMinimizeBox = (style & WS_MINIMIZEBOX) != 0;

                                    if (IsWindow(windowHandle) && IsWindowVisible(windowHandle) && !IsSystemWindow(windowHandle) && hasMinimizeBox)
                                    {
                                        foundNewWindow = true;  // 標記發現新窗口
                                        targetWindowHandle = windowHandle; // 更新為最新發現的窗口
                                        lastNewWindowTime = DateTime.Now;
                                        initialWindowHandles.Add(windowHandle); // 添加到已知集合，避免重複處理
                                    }
                                }
                                catch (Exception)
                                {
                                    // 忽略單個窗口的錯誤
                                }
                            }
                        }

                        // 只有在確實發現新窗口，且已經 2 秒沒有發現新窗口時，才執行最小化
                        if (foundNewWindow && targetWindowHandle != IntPtr.Zero && (DateTime.Now - lastNewWindowTime).TotalSeconds >= 2)
                        {
                            // 增加額外延遲，給予無回應的應用更多時間恢復
                            await Task.Delay(1000);

                            bool minimized = await ForceMinimizeWindowAsync(targetWindowHandle, "tracked window");
                            if (minimized)
                            {
                                // 最小化成功，重置狀態以繼續監控下一個可能的窗口
                                foundNewWindow = false;
                                targetWindowHandle = IntPtr.Zero;
                            }
                            // 如果最小化失敗，不要立即返回，循環將繼續，以便在下一個週期重試
                        }
                    }
                    catch (Exception)
                    {
                        // 忽略枚舉或處理過程中的錯誤
                    }
                }
                
                // 如果監控 10 秒後仍未發現新窗口，則直接返回，不進行任何最小化操作
                // 這樣可以避免對不存在的窗口進行操作
            }
            catch (Exception)
            {
                // 忽略所有異常
            }
        }

        private static List<IntPtr> EnumerateAllVisibleWindows()
        {
            var windows = new List<IntPtr>();
            try
            {
                EnumWindows((hWnd, lParam) =>
                {
                    try
                    {
                        if (IsWindowVisible(hWnd))
                        {
                            windows.Add(hWnd);
                        }
                    }
                    catch
                    {
                        // 忽略單個窗口的錯誤
                    }
                    return true;
                }, IntPtr.Zero);
            }
            catch (Exception)
            {
                // 忽略列舉窗口時的錯誤
            }
            return windows;
        }

        private static bool IsSystemWindow(IntPtr hWnd)
        {
            try
            {
                // 獲取窗口類名
                var className = new StringBuilder(256);
                GetClassName(hWnd, className, className.Capacity);
                string classNameStr = className.ToString();

                // 排除系統窗口類
                var systemClasses = new[] { "Shell_TrayWnd", "Button", "Static", "Edit", "ComboBox", "ListBox" };
                if (systemClasses.Contains(classNameStr))
                    return true;

                // 排除隱藏窗口
                if (!IsWindowVisible(hWnd))
                    return true;

                // 排除所有者窗口（通常是對話框或工具窗口）
                IntPtr owner = GetWindow(hWnd, GW_OWNER);
                if (owner != IntPtr.Zero)
                    return true;

                return false;
            }
            catch
            {
                return false;
            }
        }

        private static async Task<bool> ForceMinimizeWindowAsync(IntPtr windowHandle, string context)
        {
            if (windowHandle == IntPtr.Zero)
            {
                return false;
            }

            const int maxAttempts = 5;
            const int pollIterations = 6;
            const int pollDelayMs = 500;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                if (!IsWindow(windowHandle))
                {
                    return false;
                }
                await WaitForWindowResponsiveAsync(windowHandle, 20000); //等待視窗回應
                bool postMessageResult = PostMessage(windowHandle, WM_SYSCOMMAND, new IntPtr(SC_MINIMIZE), IntPtr.Zero);
                bool showWindowResult = ShowWindow(windowHandle, SW_MINIMIZE);

                for (int poll = 0; poll < pollIterations; poll++)
                {
                    await Task.Delay(pollDelayMs);

                    if (!IsWindow(windowHandle))
                    {
                        return true;
                    }

                    if (IsIconic(windowHandle))
                    {
                        return true;
                    }
                }

                if (attempt < maxAttempts)
                {
                    await Task.Delay(1000);
                }
            }

            return false;
        }

        private static async Task<bool> WaitForWindowResponsiveAsync(IntPtr hWnd, int timeoutMs)
        {
            if (hWnd == IntPtr.Zero)
                return false;

            var stopwatch = Stopwatch.StartNew();
            while (stopwatch.ElapsedMilliseconds < timeoutMs)
            {
                if (!IsWindow(hWnd)) // 視窗已關閉
                    return false;

                var result = SendMessageTimeout(hWnd, WM_NULL, IntPtr.Zero, IntPtr.Zero, SendMessageTimeoutFlags.SMTO_BLOCK | SendMessageTimeoutFlags.SMTO_ABORTIFHUNG, 500, out _);
                if (result != IntPtr.Zero) // 0 表示失敗或超時，非0表示成功
                {
                    return true; // 視窗有回應
                }

                // 等待一小段時間再重試
                await Task.Delay(100);
            }

            return false; // 超時，視窗仍無回應
        }

        private static async Task<IntPtr> WaitForMainWindowHandleAsync(Process process, int timeoutMs, int intervalMs)
        {
            if (process == null)
            {
                return IntPtr.Zero;
            }

            try
            {
                var stopwatch = Stopwatch.StartNew();
                while (stopwatch.ElapsedMilliseconds < timeoutMs)
                {
                    if (IsProcessExitedOrInaccessible(process))
                    {
                        return IntPtr.Zero;
                    }

                    process.Refresh();
                    if (process.MainWindowHandle != IntPtr.Zero)
                    {
                        return process.MainWindowHandle;
                    }

                    await Task.Delay(intervalMs);
                }

            }
            catch (Exception)
            {
            }

            return IntPtr.Zero;
        }

        private static void Log(string message)
        {
            try
            {
                string logFilePath = Path.Combine(AppContext.BaseDirectory, "autostarter.log");
                File.AppendAllText(logFilePath, $"{DateTime.Now:yyyy/M/d HH:mm:ss.fff}: {message}{Environment.NewLine}");
            }
            catch
            {
                // Ignore logging errors
            }
        }
    }
}
