// A native .NET apphost exercises the same process invocation as hhc.exe on Windows.
var mode = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "mode.txt")).Trim();
File.WriteAllText(Path.Combine(AppContext.BaseDirectory, "started.txt"), "started");
if (mode == "fail") return 1;

var setting = File.ReadLines(args[0]).First(line => line.StartsWith("Compiled file="));
File.WriteAllText(setting["Compiled file=".Length..], "fake chm");
return int.Parse(mode);
