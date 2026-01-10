using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media.Animation;
using Wpf.Ui.Controls;
using MessageBoxResult = Wpf.Ui.Controls.MessageBoxResult;

namespace AutoStarter;

public partial class ThemedDialogWindow : FluentWindow
{
    public string Message
    {
        get => MessageTextBlock.Text;
        set => MessageTextBlock.Text = value;
    }

    public string PrimaryButtonLabel
    {
        get => PrimaryButton.Content?.ToString() ?? string.Empty;
        set => PrimaryButton.Content = value;
    }

    public ThemedDialogWindow()
    {
        InitializeComponent();
    }

    public Task<MessageBoxResult> ShowDialogAsync()
    {
        return Dispatcher.InvokeAsync(() =>
        {
            bool? dialogResult = ShowDialog();
            return dialogResult == true ? MessageBoxResult.Primary : MessageBoxResult.None;
        }).Task;
    }

    private void PrimaryButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void SecondaryButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }

    private void DialogWindow_Loaded(object sender, RoutedEventArgs e)
    {
        BeginFadeInAnimation();
    }

    private void BeginFadeInAnimation()
    {
        var storyboard = new Storyboard();
        var opacityAnimation = new DoubleAnimation
        {
            From = 0,
            To = 1,
            Duration = TimeSpan.FromMilliseconds(300),
            EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
        };

        Storyboard.SetTarget(opacityAnimation, this);
        Storyboard.SetTargetProperty(opacityAnimation, new PropertyPath("Opacity"));
        storyboard.Children.Add(opacityAnimation);
        storyboard.Begin();
    }

    public void ConfigureSecondaryButton(string? label)
    {
        if (string.IsNullOrWhiteSpace(label))
        {
            SecondaryButton.Visibility = Visibility.Collapsed;
        }
        else
        {
            SecondaryButton.Content = label;
            SecondaryButton.Visibility = Visibility.Visible;
        }
    }
}
