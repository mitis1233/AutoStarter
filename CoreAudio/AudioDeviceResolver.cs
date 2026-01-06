using System;
using System.Collections.Generic;
using System.Linq;
using NAudio.CoreAudioApi;

namespace AutoStarter.CoreAudio
{
    internal static class AudioDeviceResolver
    {
        private static readonly StringComparer Comparer = StringComparer.OrdinalIgnoreCase;

        public static DeviceInfo? ResolveDevice(IEnumerable<DeviceInfo> devices, string? deviceId, string? instanceId, string? friendlyName)
        {
            var deviceList = devices?.ToList() ?? new List<DeviceInfo>();

            if (!string.IsNullOrWhiteSpace(deviceId))
            {
                var match = deviceList.FirstOrDefault(d => Comparer.Equals(d.ID, deviceId));
                if (match != null)
                {
                    return match;
                }
            }

            if (!string.IsNullOrWhiteSpace(instanceId))
            {
                var match = deviceList.FirstOrDefault(d => Comparer.Equals(d.InstanceId, instanceId));
                if (match != null)
                {
                    return match;
                }
            }

            if (!string.IsNullOrWhiteSpace(friendlyName))
            {
                var exactMatch = GetBestStateMatch(deviceList.Where(d => Comparer.Equals(d.FriendlyName, friendlyName)));
                if (exactMatch != null)
                {
                    return exactMatch;
                }

                var fuzzyMatch = GetBestStateMatch(deviceList.Where(d => d.FriendlyName != null && d.FriendlyName.Contains(friendlyName, StringComparison.OrdinalIgnoreCase)));
                if (fuzzyMatch != null)
                {
                    return fuzzyMatch;
                }
            }

            return null;
        }

        private static DeviceInfo? GetBestStateMatch(IEnumerable<DeviceInfo> candidates)
        {
            return candidates
                .OrderByDescending(d => GetStateScore(d.State))
                .ThenBy(d => d.FriendlyName)
                .FirstOrDefault();
        }

        private static int GetStateScore(DeviceState state) => state switch
        {
            DeviceState.Active => 4,
            DeviceState.Disabled => 3,
            DeviceState.Unplugged => 2,
            DeviceState.NotPresent => 1,
            _ => 0
        };
    }
}
