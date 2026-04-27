using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static class PsdExportManagerProgram
{
    [STAThread]
    private static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new PsdExportManagerForm());
    }
}

internal sealed class PsdExportManagerForm : Form
{
    private const string DefaultRepoUrl = "https://github.com/missingfish1234/PSD_OutPut.git";
    private const string DefaultRefName = "main";

    private readonly TextBox repoBox = new TextBox();
    private readonly TextBox refBox = new TextBox();
    private readonly Label statusLabel = new Label();
    private readonly Label installedLabel = new Label();
    private readonly Label remoteLabel = new Label();
    private readonly Button installButton = new Button();
    private readonly Button checkButton = new Button();
    private readonly Button updateButton = new Button();
    private readonly Button photoshopButton = new Button();
    private readonly Button folderButton = new Button();
    private readonly TextBox logBox = new TextBox();

    private readonly string stateRoot;
    private readonly string configPath;
    private readonly string worktreePath;
    private string lastInstalledCommit = "";
    private string lastRemoteCommit = "";

    public PsdExportManagerForm()
    {
        stateRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PSDExportPipeline");
        configPath = Path.Combine(stateRoot, "manager.ini");
        worktreePath = Path.Combine(stateRoot, "repo");

        Text = "PSD Export Manager";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(720, 520);
        Size = new Size(820, 620);
        BackColor = Color.FromArgb(43, 43, 43);
        ForeColor = Color.WhiteSmoke;

        BuildUi();
        LoadConfig();
        RefreshStateLabels();
    }

    private void BuildUi()
    {
        var root = new TableLayoutPanel();
        root.Dock = DockStyle.Fill;
        root.Padding = new Padding(14);
        root.ColumnCount = 1;
        root.RowCount = 6;
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        var title = new Label();
        title.Text = "PSD Export Pipeline";
        title.Font = new Font(Font.FontFamily, 18, FontStyle.Bold);
        title.AutoSize = true;
        title.Margin = new Padding(0, 0, 0, 10);
        root.Controls.Add(title);

        var repoPanel = new TableLayoutPanel();
        repoPanel.Dock = DockStyle.Top;
        repoPanel.ColumnCount = 2;
        repoPanel.RowCount = 2;
        repoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        repoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
        repoPanel.Margin = new Padding(0, 0, 0, 10);
        root.Controls.Add(repoPanel);

        repoPanel.Controls.Add(MakeLabel("Git Repo URL"), 0, 0);
        repoPanel.Controls.Add(MakeLabel("Branch / Tag"), 1, 0);
        repoBox.Dock = DockStyle.Fill;
        repoBox.Text = DefaultRepoUrl;
        repoPanel.Controls.Add(repoBox, 0, 1);
        refBox.Dock = DockStyle.Fill;
        refBox.Text = DefaultRefName;
        repoPanel.Controls.Add(refBox, 1, 1);

        var infoPanel = new TableLayoutPanel();
        infoPanel.Dock = DockStyle.Top;
        infoPanel.ColumnCount = 3;
        infoPanel.RowCount = 1;
        infoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
        infoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33));
        infoPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 34));
        infoPanel.Margin = new Padding(0, 0, 0, 12);
        root.Controls.Add(infoPanel);

        installedLabel.AutoSize = true;
        remoteLabel.AutoSize = true;
        statusLabel.AutoSize = true;
        infoPanel.Controls.Add(installedLabel, 0, 0);
        infoPanel.Controls.Add(remoteLabel, 1, 0);
        infoPanel.Controls.Add(statusLabel, 2, 0);

        var buttons = new FlowLayoutPanel();
        buttons.Dock = DockStyle.Top;
        buttons.AutoSize = true;
        buttons.WrapContents = true;
        buttons.Margin = new Padding(0, 0, 0, 12);
        root.Controls.Add(buttons);

        ConfigureButton(installButton, "安裝目前版本", 132, InstallBundledClicked);
        ConfigureButton(checkButton, "檢查更新", 108, CheckClicked);
        ConfigureButton(updateButton, "更新並安裝", 122, UpdateInstallClicked);
        ConfigureButton(photoshopButton, "開啟 Photoshop", 128, OpenPhotoshopClicked);
        ConfigureButton(folderButton, "開啟資料夾", 108, OpenFolderClicked);
        buttons.Controls.Add(installButton);
        buttons.Controls.Add(checkButton);
        buttons.Controls.Add(updateButton);
        buttons.Controls.Add(photoshopButton);
        buttons.Controls.Add(folderButton);

        logBox.Dock = DockStyle.Fill;
        logBox.Multiline = true;
        logBox.ScrollBars = ScrollBars.Both;
        logBox.ReadOnly = true;
        logBox.BackColor = Color.FromArgb(31, 31, 31);
        logBox.ForeColor = Color.FromArgb(230, 230, 230);
        logBox.Font = new Font("Consolas", 9);
        root.Controls.Add(logBox);

        var hint = new Label();
        hint.AutoSize = true;
        hint.Text = "流程：安裝目前版本可直接安裝旁邊的 CCX；檢查更新會比對 GitHub；更新並安裝會下載新版、重包 CCX、呼叫 Adobe 安裝器。";
        hint.ForeColor = Color.FromArgb(210, 210, 210);
        hint.Margin = new Padding(0, 10, 0, 0);
        root.Controls.Add(hint);
    }

    private static Label MakeLabel(string text)
    {
        var label = new Label();
        label.Text = text;
        label.AutoSize = true;
        label.ForeColor = Color.FromArgb(220, 220, 220);
        label.Margin = new Padding(0, 0, 0, 4);
        return label;
    }

    private static void ConfigureButton(Button button, string text, int width, EventHandler handler)
    {
        button.Text = text;
        button.Width = width;
        button.Height = 36;
        button.Margin = new Padding(0, 0, 8, 8);
        button.Click += handler;
    }

    private void LoadConfig()
    {
        Directory.CreateDirectory(stateRoot);
        if (!File.Exists(configPath))
        {
            SaveConfig();
            return;
        }

        foreach (var line in File.ReadAllLines(configPath, Encoding.UTF8))
        {
            var index = line.IndexOf('=');
            if (index <= 0) continue;
            var key = line.Substring(0, index).Trim();
            var value = line.Substring(index + 1).Trim();
            if (key.Equals("repoUrl", StringComparison.OrdinalIgnoreCase)) repoBox.Text = value;
            if (key.Equals("ref", StringComparison.OrdinalIgnoreCase)) refBox.Text = value;
            if (key.Equals("lastInstalledCommit", StringComparison.OrdinalIgnoreCase)) lastInstalledCommit = value;
        }
    }

    private void SaveConfig()
    {
        Directory.CreateDirectory(stateRoot);
        File.WriteAllLines(configPath, new[]
        {
            "repoUrl=" + RepoUrl,
            "ref=" + RefName,
            "lastInstalledCommit=" + lastInstalledCommit
        }, Encoding.UTF8);
    }

    private string RepoUrl { get { return string.IsNullOrWhiteSpace(repoBox.Text) ? DefaultRepoUrl : repoBox.Text.Trim(); } }
    private string RefName { get { return string.IsNullOrWhiteSpace(refBox.Text) ? DefaultRefName : refBox.Text.Trim(); } }
    private string AppDir { get { return Path.GetDirectoryName(Application.ExecutablePath) ?? Environment.CurrentDirectory; } }

    private void InstallBundledClicked(object sender, EventArgs e)
    {
        RunAsync("安裝目前版本", delegate
        {
            var ccx = FindLatestCcx(AppDir);
            if (string.IsNullOrEmpty(ccx))
            {
                var parentCcx = FindLatestCcx(Path.Combine(AppDir, "dist"));
                ccx = parentCcx;
            }
            if (string.IsNullOrEmpty(ccx))
            {
                throw new Exception("找不到同資料夾內的 PSDExportPipeline_*.ccx。");
            }
            InstallCcx(ccx);
            lastInstalledCommit = string.IsNullOrEmpty(lastRemoteCommit) ? lastInstalledCommit : lastRemoteCommit;
            SaveConfig();
            Log("已安裝：" + ccx);
        });
    }

    private void CheckClicked(object sender, EventArgs e)
    {
        RunAsync("檢查更新", delegate
        {
            SaveConfig();
            var remote = GetRemoteCommit();
            lastRemoteCommit = remote;
            Log("Remote commit: " + Short(remote));
            if (string.IsNullOrEmpty(lastInstalledCommit))
            {
                Log("尚未記錄已安裝 commit。可以按「更新並安裝」。");
            }
            else if (remote.Equals(lastInstalledCommit, StringComparison.OrdinalIgnoreCase))
            {
                Log("目前已是最新版本。");
            }
            else
            {
                Log("發現更新：" + Short(lastInstalledCommit) + " -> " + Short(remote));
            }
        });
    }

    private void UpdateInstallClicked(object sender, EventArgs e)
    {
        RunAsync("更新並安裝", delegate
        {
            SaveConfig();
            var remote = GetRemoteCommit();
            lastRemoteCommit = remote;
            var root = EnsureWorktree();
            UpdateWorktree(root);
            var ccx = BuildCcx(root);
            InstallCcx(ccx);
            lastInstalledCommit = GetHeadCommit(root);
            SaveConfig();
            Log("更新並安裝完成：" + Short(lastInstalledCommit));
        });
    }

    private void OpenPhotoshopClicked(object sender, EventArgs e)
    {
        RunAsync("開啟 Photoshop", delegate
        {
            var photoshop = FindPhotoshopExecutable();
            if (string.IsNullOrEmpty(photoshop))
            {
                throw new Exception("找不到 Photoshop.exe。");
            }
            Process.Start(photoshop);
            Log("已開啟：" + photoshop);
        });
    }

    private void OpenFolderClicked(object sender, EventArgs e)
    {
        RunAsync("開啟資料夾", delegate
        {
            Directory.CreateDirectory(stateRoot);
            Process.Start("explorer.exe", stateRoot);
        });
    }

    private void RunAsync(string title, Action action)
    {
        SetBusy(true);
        Log("");
        Log("== " + title + " ==");
        Task.Run(delegate
        {
            try
            {
                action();
                Ui(delegate
                {
                    RefreshStateLabels();
                    SetBusy(false);
                });
            }
            catch (Exception ex)
            {
                Log("ERROR: " + ex.Message);
                Ui(delegate
                {
                    RefreshStateLabels();
                    SetBusy(false);
                    MessageBox.Show(this, ex.Message, title + "失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        });
    }

    private void SetBusy(bool busy)
    {
        installButton.Enabled = !busy;
        checkButton.Enabled = !busy;
        updateButton.Enabled = !busy;
        photoshopButton.Enabled = !busy;
        folderButton.Enabled = !busy;
        repoBox.Enabled = !busy;
        refBox.Enabled = !busy;
        statusLabel.Text = busy ? "狀態：執行中..." : "狀態：待命";
    }

    private void RefreshStateLabels()
    {
        installedLabel.Text = "已安裝：" + (string.IsNullOrEmpty(lastInstalledCommit) ? "未記錄" : Short(lastInstalledCommit));
        remoteLabel.Text = "遠端：" + (string.IsNullOrEmpty(lastRemoteCommit) ? "尚未檢查" : Short(lastRemoteCommit));
        statusLabel.Text = "狀態：待命";
    }

    private string EnsureWorktree()
    {
        Directory.CreateDirectory(stateRoot);
        if (!Directory.Exists(Path.Combine(worktreePath, ".git")))
        {
            if (Directory.Exists(worktreePath) && Directory.EnumerateFileSystemEntries(worktreePath).Any())
            {
                throw new Exception("更新資料夾已存在但不是 Git repo：" + worktreePath);
            }
            Directory.CreateDirectory(Path.GetDirectoryName(worktreePath) ?? stateRoot);
            RunGit("clone \"" + RepoUrl + "\" \"" + worktreePath + "\"", stateRoot);
        }
        return worktreePath;
    }

    private void UpdateWorktree(string root)
    {
        RunGit("fetch --all --tags --prune", root);
        var refName = RefName;
        if (refName.StartsWith("tags/", StringComparison.OrdinalIgnoreCase))
        {
            RunGit("checkout \"refs/tags/" + refName.Substring(5) + "\"", root);
            return;
        }

        try
        {
            RunGit("checkout \"" + refName + "\"", root);
        }
        catch
        {
            RunGit("checkout -b \"" + refName + "\" \"origin/" + refName + "\"", root);
        }
        RunGit("pull --ff-only", root);
    }

    private string BuildCcx(string root)
    {
        var script = Path.Combine(root, "launcher", "build_ccx_package.ps1");
        if (!File.Exists(script))
        {
            throw new Exception("找不到 CCX 打包腳本：" + script);
        }
        RunProcess("powershell", "-NoProfile -ExecutionPolicy Bypass -File \"" + script + "\"", root);
        var ccx = FindLatestCcx(Path.Combine(root, "launcher", "dist"));
        if (string.IsNullOrEmpty(ccx))
        {
            throw new Exception("CCX 打包完成但找不到輸出檔。");
        }
        return ccx;
    }

    private void InstallCcx(string ccx)
    {
        var upia = FindUpia();
        if (string.IsNullOrEmpty(upia))
        {
            throw new Exception("找不到 UnifiedPluginInstallerAgent.exe，請安裝或更新 Adobe Creative Cloud Desktop。");
        }
        try
        {
            RunProcess(upia, "/install \"" + ccx + "\"", AppDir);
        }
        catch
        {
            RunProcess(upia, "--install \"" + ccx + "\"", AppDir);
        }
    }

    private string GetRemoteCommit()
    {
        var result = RunGit("ls-remote \"" + RepoUrl + "\" \"refs/heads/" + RefName + "\"", stateRoot);
        var first = FirstHash(result);
        if (string.IsNullOrEmpty(first))
        {
            result = RunGit("ls-remote \"" + RepoUrl + "\" HEAD", stateRoot);
            first = FirstHash(result);
        }
        if (string.IsNullOrEmpty(first))
        {
            throw new Exception("遠端 repo 沒有可用 branch/tag，請確認 GitHub 已 push 內容。");
        }
        return first;
    }

    private string GetHeadCommit(string root)
    {
        var text = RunGit("rev-parse HEAD", root);
        return text.Trim().Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
    }

    private static string FirstHash(string text)
    {
        foreach (var line in (text ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var parts = Regex.Split(line.Trim(), "\\s+");
            if (parts.Length > 0 && Regex.IsMatch(parts[0], "^[0-9a-fA-F]{7,40}$"))
            {
                return parts[0];
            }
        }
        return "";
    }

    private static string FindLatestCcx(string folder)
    {
        if (string.IsNullOrEmpty(folder) || !Directory.Exists(folder))
        {
            return "";
        }
        var files = new DirectoryInfo(folder).GetFiles("PSDExportPipeline_*.ccx");
        if (files.Length == 0)
        {
            return "";
        }
        return files.OrderByDescending(f => f.LastWriteTimeUtc).First().FullName;
    }

    private static string FindUpia()
    {
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        var candidates = new[]
        {
            Path.Combine(programFiles, "Common Files", "Adobe", "Adobe Desktop Common", "RemoteComponents", "UPI", "UnifiedPluginInstallerAgent", "UnifiedPluginInstallerAgent.exe"),
            Path.Combine(programFilesX86, "Common Files", "Adobe", "Adobe Desktop Common", "RemoteComponents", "UPI", "UnifiedPluginInstallerAgent", "UnifiedPluginInstallerAgent.exe")
        };
        return candidates.FirstOrDefault(File.Exists) ?? "";
    }

    private static string FindPhotoshopExecutable()
    {
        var adobeRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Adobe");
        if (!Directory.Exists(adobeRoot))
        {
            return "";
        }
        return new DirectoryInfo(adobeRoot)
            .GetDirectories("Adobe Photoshop*")
            .Select(d => new { Dir = d, Exe = Path.Combine(d.FullName, "Photoshop.exe"), Year = ExtractYear(d.Name), Beta = d.Name.IndexOf("beta", StringComparison.OrdinalIgnoreCase) >= 0 })
            .Where(x => File.Exists(x.Exe))
            .OrderByDescending(x => x.Beta ? x.Year - 1000 : x.Year)
            .Select(x => x.Exe)
            .FirstOrDefault() ?? "";
    }

    private static int ExtractYear(string text)
    {
        var match = Regex.Match(text ?? "", "(20\\d{2})");
        if (!match.Success) return 0;
        int year;
        return int.TryParse(match.Groups[1].Value, out year) ? year : 0;
    }

    private string RunProcess(string fileName, string arguments, string workingDirectory)
    {
        Log("> " + fileName + " " + arguments);
        var start = new ProcessStartInfo();
        start.FileName = fileName;
        start.Arguments = arguments;
        start.WorkingDirectory = string.IsNullOrEmpty(workingDirectory) ? AppDir : workingDirectory;
        start.UseShellExecute = false;
        start.RedirectStandardOutput = true;
        start.RedirectStandardError = true;
        start.CreateNoWindow = true;

        using (var process = Process.Start(start))
        {
            if (process == null)
            {
                throw new Exception("無法啟動程序：" + fileName);
            }
            var stdout = process.StandardOutput.ReadToEnd();
            var stderr = process.StandardError.ReadToEnd();
            process.WaitForExit();
            if (!string.IsNullOrWhiteSpace(stdout)) Log(stdout.TrimEnd());
            if (!string.IsNullOrWhiteSpace(stderr)) Log(stderr.TrimEnd());
            if (process.ExitCode != 0)
            {
                throw new Exception(fileName + " 結束代碼 " + process.ExitCode + "\n" + stderr.Trim());
            }
            return stdout;
        }
    }

    private string RunGit(string arguments, string workingDirectory)
    {
        var git = FindGitExecutable();
        if (string.IsNullOrEmpty(git))
        {
            throw new Exception(
                "Git was not found on this computer.\r\n" +
                "Please install Git for Windows, then reopen PSD Export Manager and try again.\r\n" +
                "Download: https://git-scm.com/download/win\r\n" +
                "If Git is already installed, make sure git.exe is available in PATH."
            );
        }
        return RunProcess(git, arguments, workingDirectory);
    }

    private static string FindGitExecutable()
    {
        var fromPath = FindExecutableInPath("git.exe");
        if (!string.IsNullOrEmpty(fromPath))
        {
            return fromPath;
        }

        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var candidates = new[]
        {
            Path.Combine(programFiles, "Git", "cmd", "git.exe"),
            Path.Combine(programFiles, "Git", "bin", "git.exe"),
            Path.Combine(programFilesX86, "Git", "cmd", "git.exe"),
            Path.Combine(programFilesX86, "Git", "bin", "git.exe"),
            Path.Combine(localAppData, "Programs", "Git", "cmd", "git.exe"),
            Path.Combine(localAppData, "Programs", "Git", "bin", "git.exe")
        };
        return candidates.FirstOrDefault(File.Exists) ?? "";
    }

    private static string FindExecutableInPath(string exeName)
    {
        var pathValue = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var rawDir in pathValue.Split(Path.PathSeparator))
        {
            var dir = (rawDir ?? "").Trim().Trim('"');
            if (string.IsNullOrEmpty(dir))
            {
                continue;
            }
            try
            {
                var candidate = Path.Combine(dir, exeName);
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            catch
            {
                // Ignore malformed PATH entries.
            }
        }
        return "";
    }

    private static string Short(string commit)
    {
        if (string.IsNullOrEmpty(commit)) return "";
        return commit.Length <= 8 ? commit : commit.Substring(0, 8);
    }

    private void Log(string text)
    {
        Ui(delegate
        {
            logBox.AppendText((text ?? "") + Environment.NewLine);
        });
    }

    private void Ui(Action action)
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            BeginInvoke(action);
            return;
        }
        action();
    }
}
