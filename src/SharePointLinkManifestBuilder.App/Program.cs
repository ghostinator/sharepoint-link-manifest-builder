using Avalonia;

namespace SharePointLinkManifestBuilder.App;

/// <summary>Application entry point.</summary>
public static class Program
{
    /// <summary>
    /// Starts the desktop application.
    /// <para>
    /// Kept free of any application logic so the Avalonia designer and the headless test host
    /// can both call <see cref="BuildAvaloniaApp"/> without starting a real UI.
    /// </para>
    /// </summary>
    /// <param name="args">Command-line arguments passed to Avalonia.</param>
    [STAThread]
    public static int Main(string[] args)
    {
        try
        {
            return BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
        }
#pragma warning disable CA1031 // Last-resort handler: report rather than vanish with no trace.
        catch (Exception ex)
        {
            // The logging pipeline may not exist yet at this point, so this writes to stderr
            // and to a crash file next to the executable rather than assuming a logger.
            Console.Error.WriteLine($"The application could not start: {ex.GetType().Name}");
            Console.Error.WriteLine(ex);

            TryWriteCrashFile(ex);
            return 1;
        }
#pragma warning restore CA1031
    }

    /// <summary>Builds the Avalonia application. Also used by the headless test host.</summary>
    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();

    private static void TryWriteCrashFile(Exception exception)
    {
        try
        {
            var path = Path.Combine(
                AppContext.BaseDirectory,
                $"startup-crash-{DateTimeOffset.UtcNow:yyyyMMdd-HHmmss}.log");

            File.WriteAllText(path, exception.ToString());
        }
#pragma warning disable CA1031 // Nothing useful remains if even this fails.
        catch (Exception)
        {
            // Deliberately ignored.
        }
#pragma warning restore CA1031
    }
}
