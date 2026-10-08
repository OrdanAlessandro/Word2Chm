using Microsoft.Win32;

namespace Word2Chm.App;

/// <summary>
/// Makes Windows open a .w2c with this application when it is double-clicked. Everything is
/// written under HKCU, so no administrator rights are needed and no other account is touched.
/// The write is best-effort: a restricted machine may refuse it, in which case projects can
/// still be opened from the toolbar.
/// </summary>
internal static class FileAssociation
{
    private const string Extension = ".w2c";
    private const string ProgId = "Word2Chm.Project";
    private const string Description = "Progetto Word2Chm";

    /// <summary>
    /// Points the .w2c extension at the running executable, doing nothing when that is
    /// already the case. Called at startup so a moved or reinstalled copy fixes itself.
    /// </summary>
    public static void EnsureRegistered()
    {
        var executable = Environment.ProcessPath;
        if (string.IsNullOrEmpty(executable))
        {
            return;
        }

        try
        {
            var command = $"\"{executable}\" \"%1\"";
            var needsRegistration =
                ReadValue($@"Software\Classes\{Extension}") != ProgId ||
                ReadValue($@"Software\Classes\{ProgId}\shell\open\command") != command ||
                ReadValue($@"Software\Classes\{ProgId}") != Description;

            if (!needsRegistration)
            {
                return;
            }

            WriteValue($@"Software\Classes\{ProgId}", Description);
            WriteValue($@"Software\Classes\{ProgId}\DefaultIcon", $"\"{executable}\",0");
            WriteValue($@"Software\Classes\{ProgId}\shell\open\command", command);

            // The extension value is what Explorer reads to pick the program; setting it last
            // means a half-written registration is never what the user sees.
            WriteValue($@"Software\Classes\{Extension}", ProgId);
        }
        catch (Exception)
        {
            // Best-effort: the association is a convenience, not a requirement.
        }
    }

    private static string? ReadValue(string keyPath)
    {
        using var key = Registry.CurrentUser.OpenSubKey(keyPath);
        return key?.GetValue(null) as string;
    }

    private static void WriteValue(string keyPath, string value)
    {
        using var key = Registry.CurrentUser.CreateSubKey(keyPath);
        key?.SetValue(null, value);
    }
}
