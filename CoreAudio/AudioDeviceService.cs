using System;
using System.Collections.Generic;
using System.Linq;
using NAudio.CoreAudioApi;

namespace AutoStarter.CoreAudio
{
    internal static class AudioDeviceService
    {
        public static (List<DeviceInfo> Playback, List<DeviceInfo> Recording) GetDeviceLists(DeviceState state)
        {
            using var enumerator = new MMDeviceEnumerator();
            return (
                EnumerateDeviceFlow(enumerator, DataFlow.Render, state),
                EnumerateDeviceFlow(enumerator, DataFlow.Capture, state));
        }

        public static List<DeviceInfo> GetAllDevices(DeviceState state)
        {
            using var enumerator = new MMDeviceEnumerator();
            var playback = EnumerateDeviceFlow(enumerator, DataFlow.Render, state);
            var recording = EnumerateDeviceFlow(enumerator, DataFlow.Capture, state);
            playback.AddRange(recording);
            return playback;
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
    }
}
