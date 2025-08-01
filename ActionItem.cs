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
    DisableAudioDevice
}

public class ActionItem : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    private bool _isEnabled = true;
    public bool IsEnabled
    {
        get => _isEnabled;
        set => SetField(ref _isEnabled, value);
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
    public string Arguments
    {
        get => _arguments;
        set => SetField(ref _arguments, value);
    }

    private int _delaySeconds;
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
    public string? AudioDeviceId
    {
        get => _audioDeviceId;
        set => SetField(ref _audioDeviceId, value);
    }

    private string? _audioDeviceName;
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

    [JsonIgnore]
    public string TypeString => Type switch
    {
        ActionType.LaunchApplication => "Application",
        ActionType.Delay => "Delay",
        ActionType.SetAudioDevice => "Audio",
        ActionType.DisableAudioDevice => "DisableAudio",
        _ => "Unknown"
    };

    [JsonIgnore]
    public string Description => Type switch
    {
        ActionType.LaunchApplication => string.IsNullOrEmpty(FilePath) ? "N/A" : Path.GetFileName(FilePath),
        ActionType.Delay => $"延遲 {DelaySeconds} 秒",
        ActionType.SetAudioDevice => string.IsNullOrEmpty(AudioDeviceName) ? "N/A" : AudioDeviceName,
        ActionType.DisableAudioDevice => string.IsNullOrEmpty(AudioDeviceName) ? "N/A" : AudioDeviceName,
        _ => "N/A"
    };

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
