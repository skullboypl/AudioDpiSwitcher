namespace AudioDpiSwitcher;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        var openSettings = args.Any(a => a.Equals("--settings", StringComparison.OrdinalIgnoreCase));
        Application.Run(new TrayApplicationContext(openSettings));
    }
}
