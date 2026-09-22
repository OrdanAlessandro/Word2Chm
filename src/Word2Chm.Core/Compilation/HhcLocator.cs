namespace Word2Chm.Core.Compilation;

/// <summary>
/// Finds hhc.exe without requiring the user to browse to it every time. The search
/// order is: an explicit user path, the HHC_PATH environment variable, then the
/// well-known installation folders of HTML Help Workshop and the Windows SDK.
/// </summary>
public static class HhcLocator
{
    public static string? Locate(string? explicitPath = null)
    {
        if (!string.IsNullOrWhiteSpace(explicitPath) && File.Exists(explicitPath))
        {
            return Path.GetFullPath(explicitPath);
        }

        var fromEnvironment = Environment.GetEnvironmentVariable("HHC_PATH");
        if (!string.IsNullOrWhiteSpace(fromEnvironment) && File.Exists(fromEnvironment))
        {
            return Path.GetFullPath(fromEnvironment);
        }

        foreach (var candidate in EnumerateCandidates())
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    private static IEnumerable<string> EnumerateCandidates()
    {
        var programFiles = new[]
        {
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
        };

        foreach (var root in programFiles.Where(r => !string.IsNullOrEmpty(r)))
        {
            yield return Path.Combine(root, "HTML Help Workshop", "hhc.exe");
            yield return Path.Combine(root, "HTML Help Workshop", "bin", "hhc.exe");

            // The Windows SDK ships hhc.exe inside versioned folders.
            var sdkRoot = Path.Combine(root, "Windows Kits", "10", "bin");
            if (Directory.Exists(sdkRoot))
            {
                foreach (var version in SafeEnumerateDirectories(sdkRoot).OrderDescending())
                {
                    foreach (var arch in new[] { "x86", "x64" })
                    {
                        yield return Path.Combine(version, arch, "hhc.exe");
                    }
                }
            }
        }

        // A copy placed next to the application still works.
        yield return Path.Combine(AppContext.BaseDirectory, "hhc.exe");
    }

    private static IEnumerable<string> SafeEnumerateDirectories(string path)
    {
        try
        {
            return Directory.EnumerateDirectories(path);
        }
        catch (IOException)
        {
            return Array.Empty<string>();
        }
        catch (UnauthorizedAccessException)
        {
            return Array.Empty<string>();
        }
    }
}
