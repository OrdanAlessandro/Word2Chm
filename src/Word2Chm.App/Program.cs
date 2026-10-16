namespace Word2Chm.App;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();

        // Registering at startup keeps the hand-off working after the executable moves; a
        // failure is ignored since the toolbar still opens projects.
        FileAssociation.EnsureRegistered();

        // Windows hands the double-clicked .w2c to the application as an argument. Opening it
        // is enough; anything else on the command line is ignored.
        var projectPath = args.Length > 0 && !string.IsNullOrWhiteSpace(args[0]) ? args[0] : null;

        Application.Run(new MainForm(projectPath));
    }
}
