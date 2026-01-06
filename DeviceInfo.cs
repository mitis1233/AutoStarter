using NAudio.CoreAudioApi;

namespace AutoStarter
{
    public class DeviceInfo
    {
        public string? ID { get; set; }
        public string? FriendlyName { get; set; }
        public string? InstanceId { get; set; }
        public DeviceState State { get; set; }

        public string DisplayName
        {
            get
            {
                var name = string.IsNullOrWhiteSpace(FriendlyName) ? "未知裝置" : FriendlyName!;
                var suffix = State switch
                {
                    DeviceState.Active => string.Empty,
                    DeviceState.Disabled => " (已停用)",
                    DeviceState.Unplugged => " (未插入)",
                    DeviceState.NotPresent => " (未連接)",
                    _ => $" ({State})"
                };
                return name + suffix;
            }
        }
    }
}
