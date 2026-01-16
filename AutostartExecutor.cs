using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using AutoStarter.CoreAudio;
using NAudio.CoreAudioApi;

namespace AutoStarter;

internal sealed partial class AutostartExecutor
{
    private static readonly JsonSerializerOptions AutostartSerializerOptions = new()
    {
        Converters = { new JsonStringEnumConverter() }
    };

    private readonly Dictionary<ActionType, Func<ActionItem, AutostartExecutionContext, Task>> _handlers;
    private IReadOnlyList<DeviceInfo>? _deviceSnapshot;
    private string? _lastLaunchedDirectory;

    public AutostartExecutor()
    {
        _handlers = new Dictionary<ActionType, Func<ActionItem, AutostartExecutionContext, Task>>
        {
            [ActionType.LaunchApplication] = HandleLaunchAsync,
            [ActionType.Delay] = HandleDelayAsync,
            [ActionType.SetAudioDevice] = HandleSetAudioDeviceAsync,
            [ActionType.DisableAudioDevice] = HandleDisableAudioDeviceAsync,
            [ActionType.EnableAudioDevice] = HandleEnableAudioDeviceAsync,
            [ActionType.SetAudioVolume] = HandleSetAudioVolumeAsync,
            [ActionType.SetPowerPlan] = HandleSetPowerPlanAsync
        };
    }

    public async Task RunAsync(string filePath)
    {
        var context = new AutostartExecutionContext();
        _deviceSnapshot = null;

        await foreach (var action in EnumerateAutostartActionsAsync(filePath))
        {
            await ExecuteActionAsync(action, context);
        }

        if (context.MinimizeTasks is { Count: > 0 })
        {
            await Task.WhenAll(context.MinimizeTasks);
        }
    }

    private Task ExecuteActionAsync(ActionItem action, AutostartExecutionContext context)
    {
        if (action == null)
        {
            return Task.CompletedTask;
        }

        return _handlers.TryGetValue(action.Type, out var handler)
            ? handler(action, context)
            : Task.CompletedTask;
    }

    private Task HandleLaunchAsync(ActionItem action, AutostartExecutionContext context)
    {
        if (!string.IsNullOrEmpty(action.FilePath) && File.Exists(action.FilePath))
        {
            var startInfo = new ProcessStartInfo(action.FilePath)
            {
                Arguments = action.Arguments,
                UseShellExecute = true,
                WorkingDirectory = GetWorkingDirectory(action.FilePath)
            };

            if (action.MinimizeWindow || action.ForceMinimizeWindow)
            {
                startInfo.WindowStyle = ProcessWindowStyle.Minimized;
            }

            var process = Process.Start(startInfo);
            _lastLaunchedDirectory = startInfo.WorkingDirectory;

            if (process != null)
            {
                if (action.ForceMinimizeWindow)
                {
                    context.MinimizeTasks ??= new List<Task>();
                    context.MinimizeTasks.Add(MinimizeWindowWithProcessMonitoringAsync(process));
                }
                else if (action.MinimizeWindow)
                {
                    context.MinimizeTasks ??= new List<Task>();
                    context.MinimizeTasks.Add(MinimizeWindowAsync(process));
                }
            }
        }

        return Task.CompletedTask;
    }

    private static Task HandleDelayAsync(ActionItem action, AutostartExecutionContext context)
    {
        return action.DelaySeconds > 0
            ? Task.Delay(TimeSpan.FromSeconds(action.DelaySeconds))
            : Task.CompletedTask;
    }

    private Task HandleSetAudioDeviceAsync(ActionItem action, AutostartExecutionContext context)
    {
        var resolvedId = ResolveAudioDeviceId(action, context);
        if (!string.IsNullOrEmpty(resolvedId))
        {
            var client = GetPolicyClient(context);
            EnableAudioDevice(client, resolvedId);
            SetDefaultAudioDevice(client, resolvedId);
        }

        return Task.CompletedTask;
    }

    private Task HandleDisableAudioDeviceAsync(ActionItem action, AutostartExecutionContext context)
    {
        var resolvedId = ResolveAudioDeviceId(action, context);
        if (!string.IsNullOrEmpty(resolvedId))
        {
            var client = GetPolicyClient(context);
            DisableAudioDevice(client, resolvedId);
        }

        return Task.CompletedTask;
    }

    private Task HandleEnableAudioDeviceAsync(ActionItem action, AutostartExecutionContext context)
    {
        var resolvedId = ResolveAudioDeviceId(action, context);
        if (!string.IsNullOrEmpty(resolvedId))
        {
            var client = GetPolicyClient(context);
            EnableAudioDevice(client, resolvedId);
        }

        return Task.CompletedTask;
    }

    private Task HandleSetAudioVolumeAsync(ActionItem action, AutostartExecutionContext context)
    {
        ApplyVolumeAdjustment(action, context);
        return Task.CompletedTask;
    }

    private static Task HandleSetPowerPlanAsync(ActionItem action, AutostartExecutionContext context)
    {
        if (action.PowerPlanId != Guid.Empty)
        {
            var guid = action.PowerPlanId;
            PowerSetActiveScheme(IntPtr.Zero, ref guid);
        }

        return Task.CompletedTask;
    }

    private static async IAsyncEnumerable<ActionItem> EnumerateAutostartActionsAsync(string filePath)
    {
        if (!File.Exists(filePath))
        {
            yield break;
        }

        await using var stream = new FileStream(
            filePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            bufferSize: 4096,
            options: FileOptions.Asynchronous | FileOptions.SequentialScan);

        await foreach (var action in JsonSerializer.DeserializeAsyncEnumerable<ActionItem>(stream, AutostartSerializerOptions))
        {
            if (action != null)
            {
                yield return action;
            }
        }
    }

    private static void SetAudioDeviceVolume(MMDeviceEnumerator enumerator, string deviceId, int volumePercent)
    {
        try
        {
            var device = enumerator.GetDevice(deviceId);
            if (device?.AudioEndpointVolume == null)
            {
                return;
            }

            var clamped = Math.Clamp(volumePercent, 0, 100) / 100f;
            device.AudioEndpointVolume.MasterVolumeLevelScalar = clamped;
        }
        catch (Exception)
        {
            // ignore
        }
    }

    private static void SetDefaultDeviceVolume(MMDeviceEnumerator enumerator, DataFlow flow, int volumePercent)
    {
        try
        {
            var device = enumerator.GetDefaultAudioEndpoint(flow, Role.Multimedia);
            if (device?.AudioEndpointVolume == null)
            {
                return;
            }

            var clamped = Math.Clamp(volumePercent, 0, 100) / 100f;
            device.AudioEndpointVolume.MasterVolumeLevelScalar = clamped;
        }
        catch (Exception)
        {
            // ignore
        }
    }

    private void ApplyVolumeAdjustment(ActionItem action, AutostartExecutionContext context)
    {
        try
        {
            using var enumerator = new MMDeviceEnumerator();
            bool applied = false;

            if (action.AdjustPlaybackVolume && action.PlaybackVolumePercent is int playbackPercent)
            {
                SetDefaultDeviceVolume(enumerator, DataFlow.Render, playbackPercent);
                applied = true;
            }

            if (action.AdjustRecordingVolume && action.RecordingVolumePercent is int recordingPercent)
            {
                SetDefaultDeviceVolume(enumerator, DataFlow.Capture, recordingPercent);
                applied = true;
            }

            if (!applied && action.AudioVolumePercent is int legacyPercent)
            {
                var resolvedId = ResolveAudioDeviceId(action, context);
                if (!string.IsNullOrEmpty(resolvedId))
                {
                    SetAudioDeviceVolume(enumerator, resolvedId, legacyPercent);
                    applied = true;
                }

                if (!applied)
                {
                    SetDefaultDeviceVolume(enumerator, DataFlow.Render, legacyPercent);
                }
            }
        }
        catch (Exception)
        {
            // ignore
        }
    }

    private static PolicyConfigClient GetPolicyClient(AutostartExecutionContext context)
    {
        context.PolicyClient ??= new PolicyConfigClient();
        return context.PolicyClient;
    }

    private static void SetDefaultAudioDevice(PolicyConfigClient client, string deviceId)
    {
        try
        {
            client.SetDefaultEndpoint(deviceId, ERole.eConsole);
            client.SetDefaultEndpoint(deviceId, ERole.eMultimedia);
            client.SetDefaultEndpoint(deviceId, ERole.eCommunications);
        }
        catch (Exception)
        {
            // ignore
        }
    }

    private static void DisableAudioDevice(PolicyConfigClient client, string deviceId)
    {
        try
        {
            client.DisableEndpoint(deviceId);
        }
        catch (Exception)
        {
            // ignore
        }
    }

    private static void EnableAudioDevice(PolicyConfigClient client, string deviceId)
    {
        try
        {
            client.EnableEndpoint(deviceId);
        }
        catch (Exception)
        {
            // ignore
        }
    }

    private string? ResolveAudioDeviceId(ActionItem action, AutostartExecutionContext context)
    {
        if (action == null)
        {
            return null;
        }

        if (string.IsNullOrWhiteSpace(action.AudioDeviceId)
            && string.IsNullOrWhiteSpace(action.AudioDeviceInstanceId)
            && string.IsNullOrWhiteSpace(action.AudioDeviceName))
        {
            return null;
        }

        var cacheKey = BuildDeviceCacheKey(action);
        if (context.ResolvedDeviceCache.TryGetValue(cacheKey, out var cached))
        {
            return ApplyResolvedDevice(action, cached);
        }

        var device = ResolveFromSnapshot(action, refresh: false);
        if (device == null && !context.DeviceSnapshotRefreshed)
        {
            context.DeviceSnapshotRefreshed = true;
            device = ResolveFromSnapshot(action, refresh: true);
        }

        context.ResolvedDeviceCache[cacheKey] = device;
        return ApplyResolvedDevice(action, device);
    }

    private static string BuildDeviceCacheKey(ActionItem action)
    {
        return string.Concat(
            action.AudioDeviceId ?? string.Empty,
            "|",
            action.AudioDeviceInstanceId ?? string.Empty,
            "|",
            action.AudioDeviceName ?? string.Empty);
    }

    private static string? ApplyResolvedDevice(ActionItem action, DeviceInfo? device)
    {
        if (device == null)
        {
            return null;
        }

        if (!string.Equals(action.AudioDeviceId, device.ID, StringComparison.OrdinalIgnoreCase))
        {
            action.AudioDeviceId = device.ID;
        }

        action.AudioDeviceInstanceId = device.InstanceId;
        action.AudioDeviceName = device.FriendlyName;

        return device.ID;
    }

    private DeviceInfo? ResolveFromSnapshot(ActionItem action, bool refresh)
    {
        try
        {
            var devices = GetDeviceSnapshot(refresh);
            return AudioDeviceResolver.ResolveDevice(
                devices,
                action.AudioDeviceId,
                action.AudioDeviceInstanceId,
                action.AudioDeviceName);
        }
        catch
        {
            return null;
        }
    }

    private IReadOnlyList<DeviceInfo> GetDeviceSnapshot(bool refresh)
    {
        if (refresh || _deviceSnapshot == null)
        {
            _deviceSnapshot = AudioDeviceService.GetAllDevices(DeviceState.All);
        }

        return _deviceSnapshot;
    }

    private string GetWorkingDirectory(string filePath)
    {
        try
        {
            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
            {
                return directory;
            }
        }
        catch
        {
            // ignored, fallback below
        }

        return _lastLaunchedDirectory ?? Environment.CurrentDirectory;
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
                        // no-op
                    }

                    // 指數退避：150ms, 300ms
                    if (attempt < 2)
                        await Task.Delay(150 * (int)Math.Pow(2, attempt));
                }
                catch (Exception) when (IsProcessExitedOrInaccessible(process))
                {
                    return;
                }
                catch (Exception ex) when (ex is ObjectDisposedException ||
                                        ex is InvalidOperationException ||
                                        ex is System.ComponentModel.Win32Exception)
                {
                    if (attempt == 2)
                        throw;
                }
            }
        }
        catch (Exception)
        {
            // ignore
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

            var initialWindowHandles = new HashSet<IntPtr>(EnumerateAllVisibleWindows());
            if (mainWindowHandle != IntPtr.Zero)
            {
                initialWindowHandles.Add(mainWindowHandle);
            }

            IntPtr targetWindowHandle = IntPtr.Zero;
            DateTime lastNewWindowTime = DateTime.MinValue;
            bool foundNewWindow = false;

            for (int i = 0; i < 20; i++)
            {
                await Task.Delay(1000);

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
                                    foundNewWindow = true;
                                    targetWindowHandle = windowHandle;
                                    lastNewWindowTime = DateTime.Now;
                                    initialWindowHandles.Add(windowHandle);
                                }
                            }
                            catch (Exception)
                            {
                                // ignore
                            }
                        }
                    }

                    if (foundNewWindow && targetWindowHandle != IntPtr.Zero && (DateTime.Now - lastNewWindowTime).TotalSeconds >= 2)
                    {
                        await Task.Delay(1000);

                        bool minimized = await ForceMinimizeWindowAsync(targetWindowHandle, "tracked window");
                        if (minimized)
                        {
                            foundNewWindow = false;
                            targetWindowHandle = IntPtr.Zero;
                        }
                    }
                }
                catch (Exception)
                {
                    // ignore
                }
            }
        }
        catch (Exception)
        {
            // ignore
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
                    // ignore
                }
                return true;
            }, IntPtr.Zero);
        }
        catch (Exception)
        {
            // ignore
        }
        return windows;
    }

    private static bool IsSystemWindow(IntPtr hWnd)
    {
        try
        {
            var className = new StringBuilder(256);
            GetClassName(hWnd, className, className.Capacity);
            string classNameStr = className.ToString();

            var systemClasses = new[] { "Shell_TrayWnd", "Button", "Static", "Edit", "ComboBox", "ListBox" };
            if (systemClasses.Contains(classNameStr))
                return true;

            if (!IsWindowVisible(hWnd))
                return true;

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
            await WaitForWindowResponsiveAsync(windowHandle, 20000);
            _ = PostMessage(windowHandle, WM_SYSCOMMAND, new IntPtr(SC_MINIMIZE), IntPtr.Zero);
            _ = ShowWindow(windowHandle, SW_MINIMIZE);

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
            if (!IsWindow(hWnd))
                return false;

            var result = SendMessageTimeout(hWnd, WM_NULL, IntPtr.Zero, IntPtr.Zero, SendMessageTimeoutFlags.SMTO_BLOCK | SendMessageTimeoutFlags.SMTO_ABORTIFHUNG, 500, out _);
            if (result != IntPtr.Zero)
            {
                return true;
            }

            await Task.Delay(100);
        }

        return false;
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
            // ignore
        }

        return IntPtr.Zero;
    }

    private sealed class AutostartExecutionContext
    {
        public List<Task>? MinimizeTasks { get; set; }
        public PolicyConfigClient? PolicyClient { get; set; }
        public bool DeviceSnapshotRefreshed { get; set; }
        public Dictionary<string, DeviceInfo?> ResolvedDeviceCache { get; } = new(StringComparer.OrdinalIgnoreCase);
    }
}
