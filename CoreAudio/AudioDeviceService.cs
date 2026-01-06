using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using NAudio.CoreAudioApi;

namespace AutoStarter.CoreAudio
{
    internal static class AudioDeviceService
    {
        public static (List<DeviceInfo> Playback, List<DeviceInfo> Recording) GetDeviceLists(DeviceState state)
        {
            if (TryGetDeviceLists(state, out var result, out var error))
            {
                return result;
            }

            if (state != DeviceState.Active && TryGetDeviceLists(DeviceState.Active, out result, out error))
            {
                return result;
            }

            throw error ?? new COMException("Unable to enumerate audio devices.");
        }

        public static List<DeviceInfo> GetAllDevices(DeviceState state)
        {
            var (playback, recording) = GetDeviceLists(state);
            playback.AddRange(recording);
            return playback;
        }

        private static bool TryGetDeviceLists(DeviceState state, out (List<DeviceInfo> Playback, List<DeviceInfo> Recording) result, out COMException? error)
        {
            var sanitizedState = SanitizeState(state);
            try
            {
                using var enumerator = new MMDeviceEnumerator();
                result = (
                    EnumerateDeviceFlow(enumerator, DataFlow.Render, sanitizedState),
                    EnumerateDeviceFlow(enumerator, DataFlow.Capture, sanitizedState));
                error = null;
                return true;
            }
            catch (COMException ex)
            {
                result = default;
                error = ex;
                return false;
            }
        }

        private static DeviceState SanitizeState(DeviceState state)
        {
            return state == DeviceState.All
                ? DeviceState.Active | DeviceState.Disabled | DeviceState.Unplugged
                : state;
        }

        private static List<DeviceInfo> EnumerateDeviceFlow(MMDeviceEnumerator enumerator, DataFlow flow, DeviceState state)
        {
            return enumerator
                .EnumerateAudioEndPoints(flow, state)
                .Select(device => new DeviceInfo
                {
                    ID = device.ID,
                    FriendlyName = device.FriendlyName,
                    InstanceId = TryGetInstanceId(device),
                    State = device.State
                })
                .OrderBy(d => GetStateSortKey(d.State))
                .ThenBy(d => d.FriendlyName)
                .ToList();
        }

        private static string? TryGetInstanceId(MMDevice device)
        {
            try
            {
                var propertyValue = device.Properties[DevicePropertyKeys.DeviceInstanceId];
                return propertyValue?.Value?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static int GetStateSortKey(DeviceState state) => state switch
        {
            DeviceState.Active => 0,
            DeviceState.Disabled => 1,
            DeviceState.Unplugged => 2,
            _ => 3
        };
    }
}
