#region Windows Task Setup And Revision 103 Service Retirement
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.ServiceProcess;
using System.Windows;
using Microsoft.Win32;
namespace GNA_DLRreport;
// Replacement for the Revision 103 service host. No service is installed.
public static class ScanSchedulerSetup
{
    public static int Configure(string stagedPath, string action)
    {
        try
        {
            if (action is not ("save" or "start" or "stop")) throw new ArgumentException(message: "Unknown scheduler action.");
            using WindowsIdentity identity = WindowsIdentity.GetCurrent();
            if (!new WindowsPrincipal(ntIdentity: identity).IsInRole(role: WindowsBuiltInRole.Administrator))
                throw new InvalidOperationException(message: "Administrator approval is required to configure the scheduled task.");
            ScanSchedulerSettings settings = ScanSchedulerStore.Read<ScanSchedulerSettings>(path: stagedPath)
                ?? throw new InvalidOperationException(message: "No staged scheduler settings.");
            if (settings.OwnerSid != identity.User?.Value) throw new InvalidOperationException(message: "Use the same Windows account to configure and run the schedule.");
            settings.Validate();
            string executable = Environment.ProcessPath ?? throw new InvalidOperationException(message: "Executable path unavailable.");
            if (!string.Equals(a: Path.GetFullPath(path: settings.ExecutablePath), b: Path.GetFullPath(path: executable), comparisonType: StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(message: "The schedule must launch this application executable.");
            settings = settings with { OwnerName = identity.Name };
            using ScanWindowsScheduler scheduler = new();
            ScanTaskSnapshot prior = scheduler.Read(sid: settings.OwnerSid);
            if (action != "stop" && prior.Running) throw new InvalidOperationException(message: "Stop the schedule and allow its current scan to finish before saving or starting again.");
            ScanSchedulerPasswordWindow? prompt = null;
            if (action == "start")
            {
                prompt = new(account: settings.OwnerName);
                if (prompt.ShowDialog() != true) return 2;
            }
            try
            {
                scheduler.Disable(sid: settings.OwnerSid);
                RetireLegacyService(executable: executable);
                if (action == "stop")
                {
                    ScanSchedulerStore.SavePrivileged(settings: settings with { Enabled = false });
                    return 0;
                }
                Directory.CreateDirectory(path: ScanSchedulerStore.Root);
                using FileStream guard = new(path: ScanSchedulerStore.WorkerLockPath, mode: FileMode.OpenOrCreate, access: FileAccess.ReadWrite, share: FileShare.None);
                ScanSchedulerStore.SavePrivileged(settings: settings with { Enabled = false });
                if (action == "save") return 0;
                try
                {
                    using System.Security.SecureString password = prompt!.Password;
                    scheduler.Register(settings: settings, password: password, utcNow: DateTime.UtcNow);
                    ScanSchedulerStore.SavePrivileged(settings: settings with { Enabled = true });
                }
                catch
                {
                    scheduler.Disable(sid: settings.OwnerSid);
                    ScanSchedulerStore.SavePrivileged(settings: settings with { Enabled = false });
                    throw;
                }
                guard.Dispose();
                if (settings.StartUtc <= DateTime.UtcNow) scheduler.RunNow(sid: settings.OwnerSid);
                return 0;
            }
            finally { prompt?.ClearPassword(); }
        }
        catch (Exception ex) { MessageBox.Show(messageBoxText: ex.Message, caption: "Scan Scheduler setup", button: MessageBoxButton.OK, icon: MessageBoxImage.Error); return 1; }
    }
    private static void RetireLegacyService(string executable)
    {
        using RegistryKey? key = Registry.LocalMachine.OpenSubKey(name: @"SYSTEM\CurrentControlSet\Services\" + ScanSchedulerStore.ServiceName, writable: false);
        if (key is null) return;
        string command = key.GetValue(name: "ImagePath") as string ?? string.Empty;
        int closingQuote = command.StartsWith(value: "\"", comparisonType: StringComparison.Ordinal) ? command.IndexOf(value: '"', startIndex: 1) : -1;
        if (closingQuote < 2 || command.Substring(startIndex: closingQuote + 1).Trim() != "--scan-service" ||
            !string.Equals(a: Path.GetFileName(path: command.Substring(startIndex: 1, length: closingQuote - 1)), b: Path.GetFileName(path: executable), comparisonType: StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(message: "A service using the old scan scheduler name has an unexpected executable. Review it in Windows Services before continuing; it has not been changed.");
        using ServiceController control = new(name: ScanSchedulerStore.ServiceName);
        if (control.Status != ServiceControllerStatus.Stopped)
        {
            if (control.Status != ServiceControllerStatus.StopPending) control.Stop();
            control.WaitForStatus(desiredStatus: ServiceControllerStatus.Stopped, timeout: TimeSpan.FromSeconds(value: 60));
        }
        ProcessStartInfo info = new(fileName: Path.Combine(path1: Environment.GetFolderPath(folder: Environment.SpecialFolder.System), path2: "sc.exe"))
            { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (string argument in new[] { "config", ScanSchedulerStore.ServiceName, "start=", "disabled" }) info.ArgumentList.Add(item: argument);
        using Process process = Process.Start(startInfo: info) ?? throw new InvalidOperationException(message: "Cannot disable the old scan service.");
        string output = process.StandardOutput.ReadToEnd(); string error = process.StandardError.ReadToEnd(); process.WaitForExit();
        if (process.ExitCode != 0) throw new InvalidOperationException(message: "Cannot disable the old scan service: " + output + error);
    }
}
#endregion
