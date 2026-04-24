using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

internal static class PsdExportLauncher
{
    private enum LauncherAction
    {
        InstallAndRun,
        InstallOnly,
        RunOnly,
        Uninstall,
        Status,
        Help,
    }

    private sealed class Options
    {
        public LauncherAction Action = LauncherAction.InstallAndRun;
        public string PluginRoot = string.Empty;
        public bool PreferBeta;
    }

    private sealed class ManifestInfo
    {
        public string PluginId = string.Empty;
        public string Version = string.Empty;
    }

    private sealed class PhotoshopCandidate
    {
        public string ExePath = string.Empty;
        public int Score;
        public int Year;
    }

    private static readonly string[] HostNames = { "PHSP", "PHSPBETA" };
    private static readonly string[] InstallBuckets = { "Developer", "External" };
    private static readonly string[] ExcludedDirectories =
    {
        ".git",
        ".vs",
        "node_modules",
        "launcher",
        "dist",
    };

    public static int Main(string[] args)
    {
        try
        {
            var options = ParseOptions(args ?? new string[0]);
            if (options.Action == LauncherAction.Help)
            {
                PrintHelp();
                return 0;
            }

            var pluginRoot = ResolvePluginRoot(options);
            if (string.IsNullOrEmpty(pluginRoot))
            {
                Console.Error.WriteLine("Could not find plugin root containing manifest.json.");
                Console.Error.WriteLine("Use --plugin-root \"F:\\path\\to\\plugin\".");
                return 2;
            }

            var manifest = ReadManifest(pluginRoot);
            if (string.IsNullOrEmpty(manifest.PluginId))
            {
                Console.Error.WriteLine("manifest.json is missing field: id");
                return 3;
            }

            Console.WriteLine("PSD Export Launcher");
            Console.WriteLine("Plugin Root : " + pluginRoot);
            Console.WriteLine("Plugin Id   : " + manifest.PluginId);
            Console.WriteLine("Version     : " + (string.IsNullOrEmpty(manifest.Version) ? "(unknown)" : manifest.Version));
            Console.WriteLine(string.Empty);

            if (options.Action == LauncherAction.Uninstall)
            {
                UninstallEverywhere(manifest.PluginId);
                return 0;
            }

            if (options.Action == LauncherAction.Status)
            {
                PrintStatus(manifest.PluginId);
                return 0;
            }

            if (options.Action == LauncherAction.InstallAndRun || options.Action == LauncherAction.InstallOnly)
            {
                InstallToAllHosts(pluginRoot, manifest);
            }

            if (options.Action == LauncherAction.InstallAndRun || options.Action == LauncherAction.RunOnly)
            {
                LaunchPhotoshop(options.PreferBeta);
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Launcher failed: " + ex.Message);
            return 1;
        }
    }

    private static Options ParseOptions(IReadOnlyList<string> args)
    {
        var options = new Options();
        for (int i = 0; i < args.Count; i++)
        {
            var arg = (args[i] ?? string.Empty).Trim();
            if (arg.Length == 0)
            {
                continue;
            }

            if (EqualsAny(arg, "-h", "--help", "/?"))
            {
                options.Action = LauncherAction.Help;
                continue;
            }

            if (EqualsAny(arg, "install"))
            {
                options.Action = LauncherAction.InstallOnly;
                continue;
            }

            if (EqualsAny(arg, "run"))
            {
                options.Action = LauncherAction.RunOnly;
                continue;
            }

            if (EqualsAny(arg, "install-run", "install+run", "all"))
            {
                options.Action = LauncherAction.InstallAndRun;
                continue;
            }

            if (EqualsAny(arg, "uninstall"))
            {
                options.Action = LauncherAction.Uninstall;
                continue;
            }

            if (EqualsAny(arg, "status"))
            {
                options.Action = LauncherAction.Status;
                continue;
            }

            if (EqualsAny(arg, "--beta"))
            {
                options.PreferBeta = true;
                continue;
            }

            if (arg.StartsWith("--plugin-root=", StringComparison.OrdinalIgnoreCase))
            {
                options.PluginRoot = NormalizePath(arg.Substring("--plugin-root=".Length));
                continue;
            }

            if (EqualsAny(arg, "--plugin-root", "-p") && i + 1 < args.Count)
            {
                i += 1;
                options.PluginRoot = NormalizePath(args[i]);
                continue;
            }

            Console.WriteLine("Unknown argument: " + arg);
            options.Action = LauncherAction.Help;
        }

        return options;
    }

    private static string ResolvePluginRoot(Options options)
    {
        var candidates = new List<string>();
        if (!string.IsNullOrWhiteSpace(options.PluginRoot))
        {
            candidates.Add(options.PluginRoot);
        }

        var exeDir = NormalizePath(AppDomain.CurrentDomain.BaseDirectory);
        candidates.Add(exeDir);
        candidates.Add(Path.GetDirectoryName(exeDir) ?? string.Empty);
        candidates.Add(Path.Combine(exeDir, "plugin"));
        var parent = Path.GetDirectoryName(exeDir);
        if (!string.IsNullOrEmpty(parent))
        {
            candidates.Add(Path.Combine(parent, "plugin"));
            var grandParent = Path.GetDirectoryName(parent);
            if (!string.IsNullOrEmpty(grandParent))
            {
                candidates.Add(grandParent);
            }
        }

        foreach (var candidate in candidates.Where((value) => !string.IsNullOrWhiteSpace(value)))
        {
            var full = NormalizePath(candidate);
            if (Directory.Exists(full) && File.Exists(Path.Combine(full, "manifest.json")))
            {
                return full;
            }
        }

        return string.Empty;
    }

    private static ManifestInfo ReadManifest(string pluginRoot)
    {
        var manifestPath = Path.Combine(pluginRoot, "manifest.json");
        var text = File.ReadAllText(manifestPath, Encoding.UTF8);
        var info = new ManifestInfo();
        info.PluginId = ExtractJsonString(text, "id");
        info.Version = ExtractJsonString(text, "version");
        return info;
    }

    private static string ExtractJsonString(string json, string field)
    {
        var pattern = "\"" + Regex.Escape(field) + "\"\\s*:\\s*\"([^\"]*)\"";
        var match = Regex.Match(json ?? string.Empty, pattern, RegexOptions.IgnoreCase);
        return match.Success ? match.Groups[1].Value : string.Empty;
    }

    private static void InstallToAllHosts(string pluginRoot, ManifestInfo manifest)
    {
        var pluginId = manifest.PluginId;
        var pluginVersion = string.IsNullOrWhiteSpace(manifest.Version) ? "0.0.0" : manifest.Version.Trim();
        var installs = new List<string>();
        foreach (var versionDir in EnumerateHostVersionDirectories())
        {
            foreach (var bucket in InstallBuckets)
            {
                var bucketDir = Path.Combine(versionDir.FullName, bucket);
                Directory.CreateDirectory(bucketDir);
                var destination = Path.Combine(bucketDir, pluginId);

                Console.WriteLine("Installing -> " + destination);
                MirrorPluginDirectory(pluginRoot, destination);
                installs.Add(destination);
            }
        }

        var localExternalRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Adobe",
            "UXP",
            "Plugins",
            "External"
        );
        Directory.CreateDirectory(localExternalRoot);
        var localExternalFolderName = $"{pluginId}_{pluginVersion}";
        var localExternalTarget = Path.Combine(localExternalRoot, localExternalFolderName);
        DeleteLocalExternalPluginCopies(localExternalRoot, pluginId);
        Console.WriteLine("Installing -> " + localExternalTarget);
        MirrorPluginDirectory(pluginRoot, localExternalTarget);
        installs.Add(localExternalTarget);

        if (installs.Count == 0)
        {
            Console.WriteLine("No PHSP/PHSPBETA version folders found in PluginsStorage.");
            Console.WriteLine("Please run Photoshop once, then run launcher again.");
            return;
        }

        Console.WriteLine(string.Empty);
        Console.WriteLine("Install completed (" + installs.Count + " targets).");
    }

    private static void UninstallEverywhere(string pluginId)
    {
        var removed = 0;
        foreach (var versionDir in EnumerateHostVersionDirectories())
        {
            foreach (var bucket in InstallBuckets)
            {
                var target = Path.Combine(versionDir.FullName, bucket, pluginId);
                if (!Directory.Exists(target))
                {
                    continue;
                }

                Console.WriteLine("Removing -> " + target);
                Directory.Delete(target, true);
                removed += 1;
            }
        }

        var localExternalRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Adobe",
            "UXP",
            "Plugins",
            "External"
        );
        if (Directory.Exists(localExternalRoot))
        {
            foreach (var dir in new DirectoryInfo(localExternalRoot).GetDirectories())
            {
                if (string.Equals(dir.Name, pluginId, StringComparison.OrdinalIgnoreCase)
                    || dir.Name.StartsWith(pluginId + "_", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Removing -> " + dir.FullName);
                    Directory.Delete(dir.FullName, true);
                    removed += 1;
                }
            }
        }

        Console.WriteLine(string.Empty);
        Console.WriteLine("Uninstall completed. Removed folders: " + removed);
    }

    private static void DeleteLocalExternalPluginCopies(string localExternalRoot, string pluginId)
    {
        if (!Directory.Exists(localExternalRoot))
        {
            return;
        }

        foreach (var dir in new DirectoryInfo(localExternalRoot).GetDirectories())
        {
            if (string.Equals(dir.Name, pluginId, StringComparison.OrdinalIgnoreCase)
                || dir.Name.StartsWith(pluginId + "_", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Removing old local external copy -> " + dir.FullName);
                Directory.Delete(dir.FullName, true);
            }
        }
    }

    private static void PrintStatus(string pluginId)
    {
        var found = new List<string>();
        foreach (var versionDir in EnumerateHostVersionDirectories())
        {
            foreach (var bucket in InstallBuckets)
            {
                var path = Path.Combine(versionDir.FullName, bucket, pluginId);
                if (Directory.Exists(path))
                {
                    found.Add(path);
                }
            }
        }

        var localExternalRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "Adobe",
            "UXP",
            "Plugins",
            "External"
        );
        if (Directory.Exists(localExternalRoot))
        {
            foreach (var dir in new DirectoryInfo(localExternalRoot).GetDirectories())
            {
                if (string.Equals(dir.Name, pluginId, StringComparison.OrdinalIgnoreCase)
                    || dir.Name.StartsWith(pluginId + "_", StringComparison.OrdinalIgnoreCase))
                {
                    found.Add(dir.FullName);
                }
            }
        }

        if (found.Count == 0)
        {
            Console.WriteLine("Plugin is not currently installed.");
            return;
        }

        Console.WriteLine("Installed locations:");
        foreach (var path in found)
        {
            Console.WriteLine("  " + path);
        }
    }

    private static IEnumerable<DirectoryInfo> EnumerateHostVersionDirectories()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        var roots = HostNames
            .Select((host) => Path.Combine(appData, "Adobe", "UXP", "PluginsStorage", host))
            .Where(Directory.Exists);

        var all = new List<DirectoryInfo>();
        foreach (var root in roots)
        {
            var rootDir = new DirectoryInfo(root);
            foreach (var dir in rootDir.GetDirectories())
            {
                if (!IsNumericFolder(dir.Name))
                {
                    continue;
                }
                all.Add(dir);
            }
        }

        return all.OrderByDescending((dir) => ParseIntegerSafe(dir.Name));
    }

    private static bool IsNumericFolder(string name)
    {
        return ParseIntegerSafe(name) > 0;
    }

    private static int ParseIntegerSafe(string value)
    {
        int parsed;
        if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed))
        {
            return parsed;
        }
        return -1;
    }

    private static void MirrorPluginDirectory(string sourceRoot, string destinationRoot)
    {
        if (Directory.Exists(destinationRoot))
        {
            Directory.Delete(destinationRoot, true);
        }

        Directory.CreateDirectory(destinationRoot);
        CopyDirectoryRecursive(
            new DirectoryInfo(sourceRoot),
            new DirectoryInfo(destinationRoot),
            sourceRoot
        );
    }

    private static void CopyDirectoryRecursive(DirectoryInfo sourceDir, DirectoryInfo destinationDir, string sourceRoot)
    {
        foreach (var file in sourceDir.GetFiles())
        {
            if (ShouldExcludeFile(file.Name))
            {
                continue;
            }

            var targetPath = Path.Combine(destinationDir.FullName, file.Name);
            file.CopyTo(targetPath, true);
        }

        foreach (var dir in sourceDir.GetDirectories())
        {
            var relative = GetRelativePath(sourceRoot, dir.FullName);
            if (ShouldExcludeDirectory(dir.Name, relative))
            {
                continue;
            }

            var targetSubdir = destinationDir.CreateSubdirectory(dir.Name);
            CopyDirectoryRecursive(dir, targetSubdir, sourceRoot);
        }
    }

    private static bool ShouldExcludeDirectory(string name, string relativePath)
    {
        if (name.StartsWith("_inspect_", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (ExcludedDirectories.Any((item) => string.Equals(name, item, StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        var normalized = (relativePath ?? string.Empty).Replace('/', '\\');
        if (normalized.StartsWith("launcher\\", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private static bool ShouldExcludeFile(string fileName)
    {
        var extension = Path.GetExtension(fileName) ?? string.Empty;
        if (extension.Equals(".ps1", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".cmd", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".bat", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".exe", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".pdb", StringComparison.OrdinalIgnoreCase) ||
            extension.Equals(".cs", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private static string GetRelativePath(string basePath, string fullPath)
    {
        var baseUri = new Uri(AppendDirectorySeparator(basePath));
        var fullUri = new Uri(fullPath);
        return Uri.UnescapeDataString(baseUri.MakeRelativeUri(fullUri).ToString());
    }

    private static string AppendDirectorySeparator(string path)
    {
        if (string.IsNullOrEmpty(path))
        {
            return path;
        }

        var separator = Path.DirectorySeparatorChar.ToString();
        return path.EndsWith(separator, StringComparison.Ordinal) ? path : path + separator;
    }

    private static void LaunchPhotoshop(bool preferBeta)
    {
        var photoshopExe = FindPhotoshopExecutable(preferBeta);
        if (string.IsNullOrEmpty(photoshopExe))
        {
            Console.WriteLine("Photoshop executable not found under Program Files\\Adobe.");
            return;
        }

        Console.WriteLine("Starting Photoshop: " + photoshopExe);
        Process.Start(new ProcessStartInfo
        {
            FileName = photoshopExe,
            UseShellExecute = true,
        });
    }

    private static string FindPhotoshopExecutable(bool preferBeta)
    {
        var adobeRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "Adobe"
        );
        if (!Directory.Exists(adobeRoot))
        {
            return string.Empty;
        }

        var candidates = new List<PhotoshopCandidate>();
        foreach (var dir in new DirectoryInfo(adobeRoot).GetDirectories("Adobe Photoshop*"))
        {
            var exePath = Path.Combine(dir.FullName, "Photoshop.exe");
            if (!File.Exists(exePath))
            {
                continue;
            }

            var isBeta = dir.Name.IndexOf("beta", StringComparison.OrdinalIgnoreCase) >= 0;
            var year = 0;
            var match = Regex.Match(dir.Name, "(20\\d{2})");
            if (match.Success)
            {
                int.TryParse(match.Groups[1].Value, out year);
            }

            var score = year;
            if (preferBeta && isBeta)
            {
                score += 10000;
            }
            if (!preferBeta && isBeta)
            {
                score -= 1000;
            }
            var item = new PhotoshopCandidate();
            item.ExePath = exePath;
            item.Score = score;
            item.Year = year;
            candidates.Add(item);
        }

        var best = candidates
            .OrderByDescending((item) => item.Score)
            .ThenByDescending((item) => item.Year)
            .FirstOrDefault();
        return best == null ? string.Empty : (best.ExePath ?? string.Empty);
    }

    private static bool EqualsAny(string value, params string[] candidates)
    {
        return candidates.Any((candidate) => string.Equals(value, candidate, StringComparison.OrdinalIgnoreCase));
    }

    private static string NormalizePath(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }
        return Path.GetFullPath(path.Trim().Trim('"'));
    }

    private static void PrintHelp()
    {
        Console.WriteLine("PSD Export Launcher");
        Console.WriteLine();
        Console.WriteLine("Usage:");
        Console.WriteLine("  PSDExportLauncher.exe install-run [--plugin-root <path>] [--beta]");
        Console.WriteLine("  PSDExportLauncher.exe install     [--plugin-root <path>]");
        Console.WriteLine("  PSDExportLauncher.exe run         [--beta]");
        Console.WriteLine("  PSDExportLauncher.exe status      [--plugin-root <path>]");
        Console.WriteLine("  PSDExportLauncher.exe uninstall   [--plugin-root <path>]");
        Console.WriteLine();
        Console.WriteLine("Default action: install-run");
    }
}
