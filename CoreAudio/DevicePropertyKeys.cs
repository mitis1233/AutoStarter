using System;
using PropertyKey = NAudio.CoreAudioApi.PropertyKey;

namespace AutoStarter.CoreAudio
{
    internal static class DevicePropertyKeys
    {
        // https://learn.microsoft.com/windows-hardware/drivers/install/devpkey-device-instanceid
        public static readonly PropertyKey DeviceInstanceId = new(
            new Guid("78C34FC8-104A-4ACA-9EA4-524D52996E57"),
            256);
    }
}
