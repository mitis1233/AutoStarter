using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace AutoStarter;

internal static class AutostartProfileService
{
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        Converters = { new JsonStringEnumConverter() },
        WriteIndented = true
    };

    public static async Task<List<ActionItem>> LoadAsync(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            return [];
        }

        var jsonString = await File.ReadAllTextAsync(filePath);
        return JsonSerializer.Deserialize<List<ActionItem>>(jsonString, SerializerOptions) ?? [];
    }

    public static async Task SaveAsync(string filePath, IEnumerable<ActionItem> items)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path is required.", nameof(filePath));
        }

        var cleanedItems = items.Select(CleanItem).ToList();
        var jsonString = JsonSerializer.Serialize(cleanedItems, SerializerOptions);
        await File.WriteAllTextAsync(filePath, jsonString);
    }

    private static ActionItem CleanItem(ActionItem item)
    {
        return new ActionItem
        {
            MinimizeWindow = item.MinimizeWindow,
            ForceMinimizeWindow = item.ForceMinimizeWindow,
            Type = item.Type,
            FilePath = item.FilePath,
            Arguments = string.IsNullOrWhiteSpace(item.Arguments) ? null : item.Arguments.Trim(),
            DelaySeconds = item.DelaySeconds,
            AudioDeviceId = item.AudioDeviceId,
            AudioDeviceInstanceId = item.AudioDeviceInstanceId,
            AudioDeviceName = item.AudioDeviceName,
            PowerPlanId = item.PowerPlanId,
            PowerPlanName = item.PowerPlanName,
            AudioVolumePercent = item.AudioVolumePercent,
            AdjustPlaybackVolume = item.AdjustPlaybackVolume,
            PlaybackVolumePercent = item.PlaybackVolumePercent,
            AdjustRecordingVolume = item.AdjustRecordingVolume,
            RecordingVolumePercent = item.RecordingVolumePercent
        };
    }
}
