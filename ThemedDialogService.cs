using System.Threading.Tasks;
using System.Windows;

namespace AutoStarter;

internal static class ThemedDialogService
{
    public static Task<Wpf.Ui.Controls.MessageBoxResult> ShowAsync(
        Window? owner,
        string title,
        string content,
        string primaryButtonText = "確定",
        string? closeButtonText = null)
    {
        var dialog = new ThemedDialogWindow
        {
            Owner = owner,
            Title = title,
            Message = content,
            PrimaryButtonLabel = primaryButtonText
        };

        dialog.ConfigureSecondaryButton(closeButtonText);

        return dialog.ShowDialogAsync();
    }
}
