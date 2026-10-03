#region Processing Interface
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
namespace GNA_DLRreport;

public partial class MainWindow
{
    #region File Location And Filename UTC Epoch
    private static string ProcessingSettingsPath => Path.Combine(path1: Environment.GetFolderPath(folder: Environment.SpecialFolder.LocalApplicationData),
        path2: "GNA Software", path3: "GNA_DLRreport", path4: "ScanProcessing.json");
    private const string EpochDisplayFormat = "yyyy-MM-dd HH:mm:ss";
    private string _processingTimeZoneId = string.Empty;
    private bool _restoringProcessingTab;
    private bool _processingOperation;
    private CancellationTokenSource? _processingCancellation;

    private void btnCancelProcessing_Click(object sender, RoutedEventArgs e) => _processingCancellation?.Cancel();

    private async Task RunProcessingOperationAsync(Func<Task> operation)
    {
        if (_scanBusy) return;
        _scanBusy = true;
        _processingOperation = true;
        List<(TabItem Tab, bool Enabled)> otherTabs = new();
        foreach (object item in tabApplication.Items)
            if (item is TabItem tab && tab != tabScan) { otherTabs.Add(item: (tab, tab.IsEnabled)); tab.IsEnabled = false; }
        UpdateScanAvailability();
        try { await operation(); }
        catch (OperationCanceledException)
        {
            gridProcessingResults.ItemsSource = null;
            txtScanStatus.Text = "Processing cancelled. No new readings or offsets saved.";
        }
        catch (Exception ex)
        {
            txtScanStatus.Text = $"Operation failed: {ex.Message}. Displayed results have not been saved.";
            MessageBox.Show(owner: this, messageBoxText: ex.Message, caption: "Scan processing", button: MessageBoxButton.OK, icon: MessageBoxImage.Error);
        }
        finally
        {
            // Keep the selected tab enabled throughout the operation and restore focus after WPF settles.
            _restoringProcessingTab = true;
            try
            {
                _scanBusy = false;
                _processingOperation = false;
                foreach (var saved in otherTabs) saved.Tab.IsEnabled = saved.Enabled;
                UpdateScanAvailability();
                pnlScanWorkspace.SelectedItem = tabScanProcessing;
                await Dispatcher.InvokeAsync(callback: () => { pnlScanWorkspace.SelectedItem = tabScanProcessing; tabScanProcessing.Focus(); },
                    priority: System.Windows.Threading.DispatcherPriority.ContextIdle);
            }
            finally { _restoringProcessingTab = false; }
        }
    }


    private void InitialiseScanProcessing()
    {
        txtProcessingEpoch.Text = string.Empty;
        try
        {
            txtProcessingFile.Text = ReadProcessingLocation(settingsPath: ProcessingSettingsPath);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException or ArgumentException or NotSupportedException)
        { txtScanStatus.Text = "Unable to restore the previous LAS location: " + ex.Message; }
    }

    private void PresentProcessingTimestamp()
    {
        if (string.IsNullOrWhiteSpace(value: txtProcessingFile.Text))
        { txtProcessingEpoch.Clear(); txtProcessingTimestampSource.Text = "Select a LAS file."; return; }
        try
        {
            ScanTimestamp stamp = TrackScan.ResolveScanTimestamp(path: txtProcessingFile.Text, timeZoneId: _processingTimeZoneId, utcNow: DateTime.UtcNow);
            txtProcessingEpoch.Text = stamp.Utc.ToString(format: EpochDisplayFormat, provider: CultureInfo.InvariantCulture);
            txtProcessingTimestampSource.Text = stamp.Source == TrackScan.FallbackTimestampSource ? "No filename timestamp: Now (UTC)" : _processingTimeZoneId;
        }
        catch (Exception ex) when (ex is InvalidOperationException or TimeZoneNotFoundException or InvalidTimeZoneException)
        { txtProcessingEpoch.Clear(); txtProcessingTimestampSource.Text = ex.Message; }
    }

    private static string ReadProcessingLocation(string settingsPath)
    {
        if (!File.Exists(path: settingsPath)) return string.Empty;
        string value = JsonSerializer.Deserialize<string>(json: File.ReadAllText(path: settingsPath)) ?? string.Empty;
        return string.IsNullOrWhiteSpace(value: value) ? string.Empty : Path.GetFullPath(path: value);
    }

    private static void SaveProcessingLocation(string settingsPath, string lasPath)
    {
        Directory.CreateDirectory(path: Path.GetDirectoryName(path: settingsPath) ?? throw new IOException(message: "No settings folder."));
        string temporary = settingsPath + "." + Guid.NewGuid().ToString(format: "N") + ".tmp";
        try
        {
            File.WriteAllText(path: temporary, contents: JsonSerializer.Serialize(value: Path.GetFullPath(path: lasPath)));
            File.Move(sourceFileName: temporary, destFileName: settingsPath, overwrite: true);
        }
        finally { if (File.Exists(path: temporary)) File.Delete(path: temporary); }
    }

    private void btnProcessingBrowse_Click(object sender, RoutedEventArgs e)
    {
        OpenFileDialog dialog = new() { Title = "Select LAS scan", Filter = "LAS scan (*.las)|*.las", CheckFileExists = true, Multiselect = false };
        if (!string.IsNullOrWhiteSpace(value: txtProcessingFile.Text))
        {
            string? directory = Path.GetDirectoryName(path: txtProcessingFile.Text);
            if (directory is not null && Directory.Exists(path: directory)) dialog.InitialDirectory = directory;
            if (File.Exists(path: txtProcessingFile.Text)) dialog.FileName = txtProcessingFile.Text;
        }
        if (dialog.ShowDialog(owner: this) != true) return;
        txtProcessingFile.Text = dialog.FileName;
        gridProcessingResults.ItemsSource = null;
        try
        {
            SaveProcessingLocation(settingsPath: ProcessingSettingsPath, lasPath: dialog.FileName);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        { txtScanStatus.Text = "LAS selected, but its location could not be remembered: " + ex.Message; }
        PresentProcessingTimestamp();
        UpdateScanAvailability();
    }
    #endregion

    #region Track Selection And Processing
    private sealed record ProcessingTrack(int TrackId, string TrackName, ScanProcessingRail Primary, ScanProcessingRail Secondary);

    private static List<ProcessingTrack> BuildProcessingTracks(IReadOnlyList<ScanProcessingRail> rails)
    {
        List<ProcessingTrack> tracks = new();
        foreach (ScanProcessingRail primary in rails)
        {
            if (primary.Rail != "R") continue;
            foreach (ScanProcessingRail secondary in rails)
            {
                if (secondary.TrackId != primary.TrackId || secondary.Rail != "L") continue;
                tracks.Add(item: new ProcessingTrack(TrackId: primary.TrackId, TrackName: primary.TrackName,
                    Primary: primary, Secondary: secondary));
                break;
            }
        }
        return tracks;
    }

    private async Task RefreshProcessingRailsAsync(ScanRepository repository, ScanProjectState state, bool clearResults = true)
    {
        int? previousTrack = (cmbProcessingTracks.SelectedItem as ProcessingTrack)?.TrackId;
        cmbProcessingTracks.ItemsSource = null;
        List<ProcessingTrack> tracks = BuildProcessingTracks(rails: await repository.LoadProcessingRailsAsync(expected: state));
        cmbProcessingTracks.ItemsSource = tracks;
        foreach (ProcessingTrack track in tracks)
            if (track.TrackId == previousTrack) cmbProcessingTracks.SelectedItem = track;
        if (cmbProcessingTracks.SelectedIndex < 0 && tracks.Count > 0) cmbProcessingTracks.SelectedIndex = 0;
        _processingTimeZoneId = await repository.LoadProcessingTimeZoneAsync();
        PresentProcessingTimestamp();
        if (clearResults) gridProcessingResults.ItemsSource = null;
        UpdateScanAvailability();
    }

    private async Task RefreshProcessingPageAsync()
    {
        txtScanStatus.Text = "Checking saved polygons for the active project...";
        ScanRepository repository = CurrentScanRepository();
        ScanProjectState state = _scanState ?? throw new InvalidOperationException(message: "Refresh Scanning first.");
        await RefreshProcessingRailsAsync(repository: repository, state: state, clearResults: false);
        txtScanStatus.Text = cmbProcessingTracks.Items.Count == 0
            ? "No tracks have current saved railhead polygons for both rails. Compute and save polygons on the Polygons tab first."
            : $"{cmbProcessingTracks.Items.Count} track(s) with current saved polygons available. Each track processes Primary then Secondary.";
    }

    private List<int> SelectedProcessingTrackIds()
    {
        List<int> selection = new();
        foreach (ProcessingTrack track in cmbProcessingTracks.Items)
            if (chkProcessingAll.IsChecked == true || ReferenceEquals(objA: track, objB: cmbProcessingTracks.SelectedItem))
                selection.Add(item: track.TrackId);
        return selection;
    }

    private static List<ScanProcessingRail> ResolveProcessingRails(IReadOnlyList<ScanProcessingRail> available, IReadOnlyList<int> selectedTracks)
    {
        List<ScanProcessingRail> rails = new();
        List<ProcessingTrack> tracks = BuildProcessingTracks(rails: available);
        foreach (int trackId in selectedTracks)
        {
            ProcessingTrack? match = null;
            foreach (ProcessingTrack track in tracks)
                if (track.TrackId == trackId) { match = track; break; }
            if (match is null) throw new InvalidOperationException(message: "Selected track polygons changed. Refresh Scanning and select the tracks again. Both rails require current polygons.");
            rails.Add(item: match.Primary);
            rails.Add(item: match.Secondary);
        }
        return rails;
    }

    private void chkProcessingAll_Click(object sender, RoutedEventArgs e)
    {
        UpdateScanAvailability();
    }

    private void cmbProcessingTracks_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (btnProcessScan is null) return;
        UpdateScanAvailability();
        e.Handled = true;
    }

    private void UpdateProcessingAvailability(bool ready)
    {
        if (btnProcessScan is null) return;
        btnSelectProcessingLas.IsEnabled = !_scanBusy;
        chkProcessingCalculateOffset.IsEnabled = !_scanBusy;
        btnCancelProcessing.Visibility = _processingCancellation is null ? Visibility.Collapsed : Visibility.Visible;
        chkProcessingAll.IsEnabled = ready && !_scanBusy && cmbProcessingTracks.Items.Count > 0;
        btnProcessScan.IsEnabled = ready && !_scanBusy && SelectedProcessingTrackIds().Count > 0 && File.Exists(path: txtProcessingFile.Text);
        cmbProcessingTracks.IsEnabled = ready && !_scanBusy && chkProcessingAll.IsChecked != true && cmbProcessingTracks.Items.Count > 0;
        txtProcessingTimestampSource.Visibility = tabScanProcessing.IsSelected ? Visibility.Visible : Visibility.Collapsed;
        if (tabScanProcessing.IsSelected)
            txtScanGuidance.Text = "Scan points are first filtered by each rail corridor, then by its railhead polygons. The LAS file is read once for all selected tracks. Results shown during processing are provisional until saved. " +
                "Filename time (YYYYMMDD_HHmmss) uses the project time zone and is stored in UTC. " +
                "With no filename timestamp, Process uses current UTC rounded to the nearest second. " +
                "Calculate Offset stores new offsets; unchecked retains and applies existing offsets. " +
                "New polygons require calibration. No accepted points means no new offset or height.";
    }

    private bool ConfirmProcessingOffsets(IReadOnlyList<ScanProcessingRail> rails, bool calculateOffsets)
    {
        string? warning = TrackScan.MissingScanCorridorMessage(rails: rails) ?? TrackScan.MissingScanOffsetMessage(rails: rails, calculateOffsets: calculateOffsets);
        if (warning is null) return true;
        txtScanStatus.Text = warning;
        MessageBox.Show(owner: this, messageBoxText: warning, caption: "Scan processing", button: MessageBoxButton.OK, icon: MessageBoxImage.Warning);
        return false;
    }

    private async void btnProcessScan_Click(object sender, RoutedEventArgs e)
    {
        List<int> selection = SelectedProcessingTrackIds();
        if (selection.Count == 0) return;
        string path = txtProcessingFile.Text;
        bool calculateOffsets = chkProcessingCalculateOffset.IsChecked == true;
        using CancellationTokenSource cancellation = new();
        _processingCancellation = cancellation;
        try
        {
            await RunProcessingOperationAsync(operation: async () =>
            {
                ScanRepository repository = CurrentScanRepository();
                ScanProjectState expected = _scanState ?? throw new InvalidOperationException(message: "Refresh Scanning first.");
                _processingTimeZoneId = await repository.LoadProcessingTimeZoneAsync();
                ScanTimestamp stamp = TrackScan.ResolveScanTimestamp(path: path, timeZoneId: _processingTimeZoneId, utcNow: DateTime.UtcNow);
                DateTime epoch = stamp.Utc;
                txtProcessingEpoch.Text = epoch.ToString(format: EpochDisplayFormat, provider: CultureInfo.InvariantCulture);
                txtProcessingTimestampSource.Text = stamp.Source == TrackScan.FallbackTimestampSource ? "No filename timestamp: Now (UTC)" : _processingTimeZoneId;
                IReadOnlyList<ScanProcessingRail> available = await repository.LoadProcessingRailsAsync(expected: expected);
                List<ScanProcessingRail> rails = ResolveProcessingRails(available: available, selectedTracks: selection);
                if (!ConfirmProcessingOffsets(rails: rails, calculateOffsets: calculateOffsets)) return;
                var pending = new System.Collections.ObjectModel.ObservableCollection<ScanEpochReading>();
                gridProcessingResults.ItemsSource = pending;
                txtScanStatus.Text = "Filtering rail corridors, then railhead polygons, Primary before Secondary...";
                Progress<ScanEpochReading> readingProgress = new(handler: reading => pending.Add(item: reading));
                Progress<string> progress = new(handler: text => txtScanStatus.Text = text);
                ScanProcessingResult result = await Task.Run(function: () => TrackScan.ProcessLas(path: path, utcTimestamp: epoch, rails: rails, progress: progress, calculateOffsets: calculateOffsets, readingProgress: readingProgress, cancellationToken: cancellation.Token));
                result = result with { TimeZoneId = _processingTimeZoneId, TimestampSource = stamp.Source };
                cancellation.Token.ThrowIfCancellationRequested();
                btnCancelProcessing.IsEnabled = false;
                txtScanStatus.Text = "Saving scan epoch readings...";
                await repository.SaveScanEpochAsync(expected: expected, result: result);
                gridProcessingResults.ItemsSource = result.Readings;
                int empty = result.Readings.Count(predicate: reading => reading.ScanPointCount == 0);
                txtScanStatus.Text = $"Saved {result.Readings.Count:N0} railhead readings to DBTrackScan at {epoch:yyyy-MM-dd HH:mm:ss} UTC. " +
                    $"{empty:N0} readings have no accepted scan points. LAS points examined: {result.LasPointCount:N0}.";
                foreach (ScanCorridorSelection corridor in result.CorridorSelections)
                    txtScanStatus.Text += Environment.NewLine + $"{corridor.TrackName} — {(corridor.Rail == "R" ? "Primary" : "Secondary")}: {corridor.ScanPointCount:N0} scan points inside corridor.";
                if (calculateOffsets) new ScanOffsetReviewWindow(result: result) { Owner = this }.ShowDialog();
            });
        }
        finally { _processingCancellation = null; btnCancelProcessing.IsEnabled = true; UpdateScanAvailability(); }
        // Returning from a warning/review is not a new request to reload and replace its status.
        _restoringProcessingTab = true;
        try { pnlScanWorkspace.SelectedItem = tabScanProcessing; }
        finally { _restoringProcessingTab = false; }
    }
    #endregion
}
#endregion
