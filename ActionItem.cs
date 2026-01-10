using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.Json.Serialization;

namespace AutoStarter;

public enum ActionType
{
    LaunchApplication,
    SetAudioDevice,
    Delay,
    DisableAudioDevice,
    EnableAudioDevice,
    SetAudioVolume,
    SetPowerPlan
}

public class ActionItem : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    private bool _minimizeWindow = false;
    public bool MinimizeWindow
    {
        get => _minimizeWindow;
        set
        {
            if (SetField(ref _minimizeWindow, value))
            {
                // 如果啟用最小化，則禁用強制最小化
                if (value && _forceMinimizeWindow)
                {
                    ForceMinimizeWindow = false;
                }
            }
        }
    }

    private int? _audioVolumePercent;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int? AudioVolumePercent
    {
        get => _audioVolumePercent;
        set
        {
            if (SetField(ref _audioVolumePercent, value))
            {
                OnPropertyChanged(nameof(Description));
            }
        }
    }

    private bool _adjustPlaybackVolume;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool AdjustPlaybackVolume
    {
        get => _adjustPlaybackVolume;
        set
        {
            if (SetField(ref _adjustPlaybackVolume, value))
            {
                OnPropertyChanged(nameof(Description));
            }
        }
    }

    private bool _adjustRecordingVolume;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public bool AdjustRecordingVolume
    {
        get => _adjustRecordingVolume;
        set
        {
            if (SetField(ref _adjustRecordingVolume, value))
            {
                OnPropertyChanged(nameof(Description));
            }
        }
    }

    private int? _playbackVolumePercent;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int? PlaybackVolumePercent
    {
        get => _playbackVolumePercent;
        set
        {
            if (SetField(ref _playbackVolumePercent, value))
            {
                OnPropertyChanged(nameof(Description));
            }
        }
    }

    private int? _recordingVolumePercent;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int? RecordingVolumePercent
    {
        get => _recordingVolumePercent;
        set
        {
            if (SetField(ref _recordingVolumePercent, value))
            {
                OnPropertyChanged(nameof(Description));
            }
        }
    }

    private bool _forceMinimizeWindow = false;
    public bool ForceMinimizeWindow
    {
        get => _forceMinimizeWindow;
        set
        {
            if (SetField(ref _forceMinimizeWindow, value))
            {
                // 如果啟用強制最小化，則禁用普通最小化
                if (value && _minimizeWindow)
                {
                    MinimizeWindow = false;
                }
            }
        }
    }

    private ActionType _type;
    public ActionType Type
    {
        get => _type;
        set
        {
            if (SetField(ref _type, value))
            {
                OnPropertyChanged(nameof(Description));
                OnPropertyChanged(nameof(TypeString));
            }
        }
    }

    private string _filePath = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string FilePath
    {
        get => _filePath;
        set
        {
            if (SetField(ref _filePath, value))
            {
                OnPropertyChanged(nameof(Description));
            }
        }
    }

    private string _arguments = string.Empty;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string Arguments
    {
        get => _arguments;
        set => SetField(ref _arguments, value);
    }

    private int _delaySeconds;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public int DelaySeconds
    {
        get => _delaySeconds;
        set
        {
            if (SetField(ref _delaySeconds, value))
            {
                OnPropertyChanged(nameof(Description));
            }
        }
    }

    private string? _audioDeviceId;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string? AudioDeviceId
    {
        get => _audioDeviceId;
        set => SetField(ref _audioDeviceId, value);
    }

    private string? _audioDeviceInstanceId;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string? AudioDeviceInstanceId
    {
        get => _audioDeviceInstanceId;
        set => SetField(ref _audioDeviceInstanceId, value);
    }

    private string? _audioDeviceName;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string? AudioDeviceName
    {
        get => _audioDeviceName;
        set
        {
            if (SetField(ref _audioDeviceName, value))
            {
                OnPropertyChanged(nameof(Description));
            }
        }
    }

    private Guid _powerPlanId;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public Guid PowerPlanId
    {
        get => _powerPlanId;
        set => SetField(ref _powerPlanId, value);
    }

    private string? _powerPlanName;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingDefault)]
    public string? PowerPlanName
    {
        get => _powerPlanName;
        set
        {
            if (SetField(ref _powerPlanName, value))
            {
                OnPropertyChanged(nameof(Description));
            }
        }
    }

    [JsonIgnore]
    public string TypeString => Type switch
    {
        ActionType.LaunchApplication => "Application",
        ActionType.Delay => "Delay",
        ActionType.SetAudioDevice => "Audio",
        ActionType.DisableAudioDevice => "DisableAudio",
        ActionType.EnableAudioDevice => "EnableAudio",
        ActionType.SetAudioVolume => "AudioVolume",
        ActionType.SetPowerPlan => "PowerPlan",
        _ => "Unknown"
    };

    [JsonIgnore]
    public string Description => Type switch
    {
        ActionType.LaunchApplication => string.IsNullOrEmpty(FilePath) ? "N/A" : Path.GetFileName(FilePath),
        ActionType.Delay => $"延遲 {DelaySeconds} 秒",
        ActionType.SetAudioDevice => string.IsNullOrEmpty(AudioDeviceName) ? "N/A" : AudioDeviceName,
        ActionType.DisableAudioDevice => string.IsNullOrEmpty(AudioDeviceName) ? "N/A" : AudioDeviceName,
        ActionType.EnableAudioDevice => string.IsNullOrEmpty(AudioDeviceName) ? "N/A" : AudioDeviceName,
        ActionType.SetAudioVolume => BuildAudioVolumeDescription(),
        ActionType.SetPowerPlan => string.IsNullOrEmpty(PowerPlanName) ? "N/A" : PowerPlanName,
        _ => "N/A"
    };

    private string BuildAudioVolumeDescription()
    {
        var parts = new List<string>();
        if (AdjustPlaybackVolume && PlaybackVolumePercent.HasValue)
        {
            parts.Add($"播放 {PlaybackVolumePercent.Value}%");
        }

        if (AdjustRecordingVolume && RecordingVolumePercent.HasValue)
        {
            parts.Add($"錄音 {RecordingVolumePercent.Value}%");
        }

        if (parts.Count > 0)
        {
            return string.Join(" / ", parts);
        }

        if (AudioVolumePercent.HasValue)
        {
            return $"音量 {AudioVolumePercent.Value}%";
        }

        if (!string.IsNullOrEmpty(AudioDeviceName))
        {
            return AudioDeviceName;
        }

        return "音量控制";
    }

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    protected bool SetField<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return false;
        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }
}
