using System;
using System.Windows;

namespace AutoStarter
{
    public partial class App : Application
    {
        protected override async void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            if (e.Args.Length > 0 && e.Args[0].EndsWith(".autostart", StringComparison.OrdinalIgnoreCase))
            {
                string filePath = e.Args[0];
                //Log($"找到 .autostart 檔案：{filePath}");
                try
                {
                    var executor = new AutostartExecutor();
                    await executor.RunAsync(filePath);
                }
                catch (Exception)//ex
                {
                    //Log($"處理 .autostart 檔案時發生錯誤：{ex.Message}");
                }

                //Log("已完成處理 .autostart 檔案。正在關閉程式。");
                Shutdown();
                return;
            }

            var mainWindow = new MainWindow();
            mainWindow.Show();
        }

    }
}
