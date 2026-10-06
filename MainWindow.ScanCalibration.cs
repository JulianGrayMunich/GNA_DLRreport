#region Calibration File Queue And Review Workflow
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;
namespace GNA_DLRreport;
public partial class MainWindow
{
    private sealed class CalibrationFile : INotifyPropertyChanged
    {
        public string Path { get; }
        public string FileName => System.IO.Path.GetFileName(path: Path);
        private string _status = "Queued", _timestamp = "Resolved when processed";
        public string Status { get => _status; set { _status = value; PropertyChanged?.Invoke(sender: this, e: new(propertyName: nameof(Status))); } }
        public string Timestamp { get => _timestamp; set { _timestamp = value; PropertyChanged?.Invoke(sender: this, e: new(propertyName: nameof(Timestamp))); } }
        public CalibrationFile(string path) { Path = path; }
        public event PropertyChangedEventHandler? PropertyChanged;
    }
    private readonly ObservableCollection<CalibrationFile> _calibrationFiles = new();
    private void AddCalibrationFiles(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            string full = Path.GetFullPath(path: path);
            if (_calibrationFiles.Any(predicate: file => string.Equals(a: file.Path, b: full, comparisonType: StringComparison.OrdinalIgnoreCase))) continue;
            _calibrationFiles.Add(item: new(path: full));
        }
        gridCalibrationFiles.ItemsSource = _calibrationFiles;
        PresentCalibrationFileTimes();
    }
    private void PresentCalibrationFileTimes()
    {
        foreach (CalibrationFile file in _calibrationFiles)
        {
            try
            {
                ScanTimestamp stamp = TrackScan.ResolveScanTimestamp(path: file.Path, timeZoneId: _processingTimeZoneId, utcNow: DateTime.UtcNow);
                file.Timestamp = stamp.Source == TrackScan.FallbackTimestampSource ? "Now (UTC) at processing" : stamp.Utc.ToString(format: EpochDisplayFormat, provider: System.Globalization.CultureInfo.InvariantCulture);
            }
            catch (Exception ex) when (ex is InvalidOperationException or TimeZoneNotFoundException or InvalidTimeZoneException)
            { file.Timestamp = ex.Message; }
        }
    }
    private void CalibrationMode_Changed(object sender, RoutedEventArgs e)
    {
        if (gridCalibrationFiles is null || txtProcessingFile is null) return;
        if (chkProcessingCalculateOffset.IsChecked == true && _calibrationFiles.Count == 0 && File.Exists(path: txtProcessingFile.Text))
            AddCalibrationFiles(paths: new[] { txtProcessingFile.Text });
        UpdateScanAvailability();
    }
    private void CalibrationRemove_Click(object sender, RoutedEventArgs e)
    { foreach (CalibrationFile file in gridCalibrationFiles.SelectedItems.Cast<CalibrationFile>().ToArray()) _calibrationFiles.Remove(item: file); UpdateScanAvailability(); }
    private void CalibrationClear_Click(object sender, RoutedEventArgs e) { _calibrationFiles.Clear(); UpdateScanAvailability(); }
    private void UpdateCalibrationFileAvailability()
    {
        if (pnlCalibrationFiles is null) return;
        bool calibrating = chkProcessingCalculateOffset.IsChecked == true;
        pnlCalibrationFiles.Visibility = calibrating ? Visibility.Visible : Visibility.Collapsed;
        btnSelectProcessingLas.Content = calibrating ? "Select LAS files" : "Select LAS file";
        btnCalibrationAdd.IsEnabled = !_scanBusy; btnCalibrationRemove.IsEnabled = !_scanBusy; btnCalibrationClear.IsEnabled = !_scanBusy;
        cmbProcessingRejectionCount.IsEnabled = !_scanBusy;
    }
    private async Task ProcessCalibrationFilesAsync(ScanRepository repository, ScanProjectState expected,
        IReadOnlyList<ScanProcessingRail> rails, CancellationTokenSource cancellation, bool debugHistograms, int rejectionCount)
    {
        CalibrationFile[] files = _calibrationFiles.ToArray();
        if (files.Length == 0) throw new InvalidOperationException(message: "Select calibration LAS files first.");
        List<ScanProcessingResult> results = new();
        HashSet<string> hashes = new(comparer: StringComparer.OrdinalIgnoreCase);
        foreach (CalibrationFile file in files) file.Status = "Queued";
        try
        {
            for (int i = 0; i < files.Length; i++)
            {
                cancellation.Token.ThrowIfCancellationRequested();
                CalibrationFile file = files[i]; file.Status = "Processing";
                ScanTimestamp stamp = TrackScan.ResolveScanTimestamp(path: file.Path, timeZoneId: _processingTimeZoneId, utcNow: DateTime.UtcNow);
                file.Timestamp = stamp.Utc.ToString(format: EpochDisplayFormat, provider: System.Globalization.CultureInfo.InvariantCulture);
                txtProcessingEpoch.Text = file.Timestamp;
                int number = i + 1;
                Progress<string> progress = new(handler: text => txtScanStatus.Text = $"File {number}/{files.Length}: {file.FileName}. {text}");
                ScanProcessingResult result = await Task.Run(function: () => TrackScan.ProcessLas(path: file.Path, utcTimestamp: stamp.Utc,
                    rails: rails, calculateOffsets: true, progress: progress, cancellationToken: cancellation.Token,
                    collectDebugHistograms: debugHistograms, rejectionCount: rejectionCount));
                result = result with { TimeZoneId = _processingTimeZoneId, TimestampSource = stamp.Source };
                if (!hashes.Add(item: result.Sha256)) throw new InvalidOperationException(message: $"'{file.FileName}' duplicates another selected scan's contents. Remove the duplicate before calibration. No offsets saved.");
                foreach (ScanHeightHistogram histogram in result.DebugHistograms)
                {
                    cancellation.Token.ThrowIfCancellationRequested();
                    ScanHeightHistogramWindow window = new(histogram: histogram) { Owner = this };
                    window.Title += " — " + file.FileName;
                    if (window.ShowDialog() != true) { cancellation.Cancel(); cancellation.Token.ThrowIfCancellationRequested(); }
                }
                results.Add(item: result with { DebugHistograms = Array.Empty<ScanHeightHistogram>() });
                file.Status = result.Warnings.Count == 0 ? "Ready for review" : $"Ready; {result.Warnings.Count} missing results";
            }
            cancellation.Token.ThrowIfCancellationRequested();
            ScanCalibrationReview review = new(results: results.AsReadOnly());
            txtScanStatus.Text = "Review calibration observations. No monitoring offsets have changed.";
            if (new ScanCalibrationReviewWindow(review: review) { Owner = this }.ShowDialog() != true)
            { cancellation.Cancel(); cancellation.Token.ThrowIfCancellationRequested(); }
            cancellation.Token.ThrowIfCancellationRequested();
            btnCancelProcessing.IsEnabled = false;
            txtScanStatus.Text = "Saving calibration observations and adopting approved mean offsets...";
            await repository.AdoptCalibrationAsync(expected: expected, review: review);
            foreach (CalibrationFile file in files) file.Status = "Calibration review saved";
            gridProcessingResults.ItemsSource = null;
            int adopted = review.Rows.Count(predicate: row => row.Mean.HasValue);
            txtScanStatus.Text = $"Adopted mean offsets for {adopted} railhead points from {files.Length} scans. {review.Rows.Count - adopted} points had no included observations and retained their existing offsets, if any. Calibration observations and decisions saved to DBTrackGeometry. Untick Calculate Offset for monitoring.";
        }
        catch
        {
            foreach (CalibrationFile file in files) if (file.Status != "Calibration review saved") file.Status = "Not adopted";
            throw;
        }
    }
}
#endregion
