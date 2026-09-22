using System.Text;
using Word2Chm.Core.Generation;
using Word2Chm.Core.Model;

namespace Word2Chm.Core.Compilation;

public sealed class CompileOptions
{
    /// <summary>Full path to hhc.exe. When null, only the project files are generated.</summary>
    public string? HhcPath { get; set; }

    public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(5);
}

public sealed class CompileResult
{
    public bool Success { get; init; }
    public int ExitCode { get; init; }
    public string Output { get; init; } = string.Empty;
    public string? ChmPath { get; init; }
    public string WorkingDirectory { get; init; } = string.Empty;
}

/// <summary>
/// Invokes HTML Help Workshop (hhc.exe) on a generated .hhp file. Compilation is
/// Windows-only in practice, but the process handling is platform-neutral so that
/// a compatible compiler (for example chmcmd) can be substituted.
/// </summary>
public sealed class ChmCompiler
{
    public CompileResult Compile(string hhpPath, string chmPath, CompileOptions options)
    {
        var workingDirectory = Path.GetDirectoryName(Path.GetFullPath(hhpPath))!;

        if (string.IsNullOrWhiteSpace(options.HhcPath))
        {
            return new CompileResult
            {
                Success = false,
                ExitCode = -1,
                Output = "Percorso di hhc.exe non configurato: i file di progetto sono stati generati ma il CHM non è stato compilato.",
                WorkingDirectory = workingDirectory,
            };
        }

        if (!File.Exists(options.HhcPath))
        {
            return new CompileResult
            {
                Success = false,
                ExitCode = -1,
                Output = $"hhc.exe non trovato: {options.HhcPath}",
                WorkingDirectory = workingDirectory,
            };
        }

        var output = new StringBuilder();
        using var process = new System.Diagnostics.Process();
        process.StartInfo = new System.Diagnostics.ProcessStartInfo
        {
            FileName = options.HhcPath,
            Arguments = "\"" + Path.GetFileName(hhpPath) + "\"",
            WorkingDirectory = workingDirectory,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };

        process.OutputDataReceived += (_, e) =>
        {
            if (e.Data is not null)
            {
                lock (output)
                {
                    output.AppendLine(e.Data);
                }
            }
        };

        process.ErrorDataReceived += (_, e) =>
        {
            if (e.Data is not null)
            {
                lock (output)
                {
                    output.AppendLine("[stderr] " + e.Data);
                }
            }
        };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();

        if (!process.WaitForExit((int)options.Timeout.TotalMilliseconds))
        {
            try
            {
                process.Kill(entireProcessTree: true);
            }
            catch (InvalidOperationException)
            {
                // The process already exited between the timeout and the kill attempt.
            }

            return new CompileResult
            {
                Success = false,
                ExitCode = -2,
                Output = output + "\nCompilazione interrotta per timeout.",
                WorkingDirectory = workingDirectory,
            };
        }

        // Ensure asynchronous output draining has completed.
        process.WaitForExit();

        var compiled = File.Exists(chmPath);
        return new CompileResult
        {
            Success = process.ExitCode == 0 && compiled,
            ExitCode = process.ExitCode,
            Output = output.ToString().Trim(),
            ChmPath = compiled ? chmPath : null,
            WorkingDirectory = workingDirectory,
        };
    }
}
