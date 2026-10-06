#region One Scheduled Scan Search
using System.IO;
using System.Security.Principal;
namespace GNA_DLRreport;
public static class ScanScheduledRunner
{
    public static async Task<int> RunAsync()
    {
        ScanSchedulerSettings? settings = null;
        DateTime started = DateTime.UtcNow;
        FileStream? guard = null;
        try
        {
            settings = ScanSchedulerStore.Read<ScanSchedulerSettings>(path: ScanSchedulerStore.PrivateConfigPath)
                ?? throw new InvalidOperationException(message: "No scan schedule is configured.");
            if (settings.OwnerSid != WindowsIdentity.GetCurrent().User?.Value) throw new InvalidOperationException(message: "The scheduled Windows account does not match the saved configuration.");
            settings.Validate();
            using (ScanWindowsScheduler scheduler = new())
                if (!settings.Enabled || !scheduler.Read(sid: settings.OwnerSid).Enabled) return 0;
            try { guard = new(path: ScanSchedulerStore.WorkerLockPath, mode: FileMode.OpenOrCreate, access: FileAccess.ReadWrite, share: FileShare.None); }
            catch (IOException) { return 3; }
            ScanSchedulerSettings? current = ScanSchedulerStore.Read<ScanSchedulerSettings>(path: ScanSchedulerStore.PrivateConfigPath);
            if (current?.Revision != settings.Revision || !current.Enabled) return 0;
            string message = "Searching incoming LAS folder";
            using CancellationTokenSource heartbeatStop = new();
            Task heartbeat = Task.Run(function: async () =>
            {
                try
                {
                    while (!heartbeatStop.IsCancellationRequested)
                    {
                        WriteStatus(settings: settings, state: "Running", message: Volatile.Read(location: ref message), started: started);
                        await Task.Delay(delay: TimeSpan.FromSeconds(value: 1), cancellationToken: heartbeatStop.Token);
                    }
                }
                catch (OperationCanceledException) { }
            });
            try
            {
                Progress<string> progress = new(handler: value => Volatile.Write(location: ref message, value: value));
                await new ScanSchedulerEngine().RunCycleAsync(settings: settings, progress: progress, cancellation: CancellationToken.None);
            }
            finally { heartbeatStop.Cancel(); await heartbeat; }
            WriteStatus(settings: settings, state: "Completed", message: "Search completed. See the system log for individual file outcomes.", started: started);
            return 0;
        }
        catch (Exception ex)
        {
            if (settings is not null)
            {
                try { ScanSchedulerEngine.Log(settings: settings, message: "Unsuccessful scheduled search: " + ex.Message); } catch { }
                WriteStatus(settings: settings, state: "Error", message: "Search failed. Check the system log and account access.", started: started);
            }
            return 1;
        }
        finally { guard?.Dispose(); }
    }
    private static void WriteStatus(ScanSchedulerSettings settings, string state, string message, DateTime started)
    {
        try { ScanSchedulerStore.Write(path: ScanSchedulerStore.StatusPath, value: new ScanSchedulerStatus(HeartbeatUtc: DateTime.UtcNow,
            Configuration: settings.Revision, State: state, Message: message, NextUtc: null, LastRunUtc: started)); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
#endregion
