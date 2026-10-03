#region System Preparation
using System.Windows;
#endregion

namespace GNA_DLRreport;

public partial class MainWindow
{
    #region Endpoint Computation And Review Before Saving
    private async void btnScanSaveTrack_Click(object sender, RoutedEventArgs e)
    {
        ScanTrack original;
        try
        {
            original = ReadScanTrack();
            ScanCoordinates.ValidateTrack(track: original);
        }
        catch (Exception ex) when (ex is FormatException or ArgumentException or InvalidOperationException)
        {
            txtScanStatus.Text = ex.Message;
            MessageBox.Show(owner: this, messageBoxText: ex.Message, caption: "Endpoint check", button: MessageBoxButton.OK, icon: MessageBoxImage.Warning);
            return;
        }

        await RunScanOperationAsync(operation: async () =>
        {
            ScanRepository repository = CurrentScanRepository();
            ScanProjectState expected = _scanState ?? throw new InvalidOperationException(message: "Refresh Scanning first.");
            txtScanStatus.Text = "Loading the current Railhead survey for the endpoint checks...";
            IReadOnlyList<ScanSurveyPoint> survey = await repository.LoadStartPointSurveyAsync(expected: expected);
            ScanStartPointCheck start = await Task.Run(function: () => TrackScan.CheckStartLeftPoint(
                rightStart: new(Easting: original.RightStartE, Northing: original.RightStartN),
                primaryEnd: new(Easting: original.RightEndE, Northing: original.RightEndN),
                suppliedLeftStart: new(Easting: original.LeftStartE, Northing: original.LeftStartN),
                pointSpacing: original.PointSpacing, gaugeMillimetres: original.GaugeMillimetres, survey: survey));
            txtScanStatus.Clear();
            if (!start.Succeeded)
            {
                MessageBox.Show(owner: this, messageBoxText: start.Warning, caption: "Endpoint computation stopped", button: MessageBoxButton.OK, icon: MessageBoxImage.Warning);
                return;
            }
            ScanEndPointCheck end = await Task.Run(function: () => TrackScan.CheckSecondaryEndPoint(
                primaryEnd: new(Easting: original.RightEndE, Northing: original.RightEndN),
                primaryStart: new(Easting: original.RightStartE, Northing: original.RightStartN),
                suppliedSecondaryEnd: new(Easting: original.LeftEndE, Northing: original.LeftEndN),
                pointSpacing: original.PointSpacing, gaugeMillimetres: original.GaugeMillimetres, survey: survey));
            txtScanStatus.Clear();
            if (!end.Succeeded)
            {
                MessageBox.Show(owner: this, messageBoxText: end.Warning, caption: "Endpoint computation stopped", button: MessageBoxButton.OK, icon: MessageBoxImage.Warning);
                return;
            }
            TrackScan.RailCoordinate secondaryStart = start.LeftStart ?? throw new InvalidOperationException(message: "Missing checked secondary start.");
            TrackScan.RailCoordinate secondaryEnd = end.SecondaryEnd ?? throw new InvalidOperationException(message: "Missing checked secondary end.");
            ScanTrack reviewed = TrackScan.RoundTrackCoordinates(track: original with
            {
                LeftStartE = secondaryStart.Easting, LeftStartN = secondaryStart.Northing,
                LeftEndE = secondaryEnd.Easting, LeftEndN = secondaryEnd.Northing
            });
            ScanCoordinates.ValidateTrack(track: reviewed, roundedCoordinates: true);
            ScanEndpointReviewWindow review = new(original: original, proposed: reviewed) { Owner = this };
            if (review.ShowDialog() != true)
            {
                txtScanStatus.Text = "Review cancelled. No coordinates have been changed or saved.";
                return;
            }

            // The repository saves the reviewed coordinates and checks the project
            // revision again to reject any intervening survey or track change.
            await repository.SaveReviewedTrackAsync(expected: expected, track: reviewed);
            _scanEditorDirty = false;
            FillScanEditor(track: reviewed);
            string status = $"Track '{reviewed.Name}' saved with the accepted endpoint coordinates.";
            if (original.TrackId != 0)
                foreach (ScanTrack previous in expected.Tracks)
                    if (previous.TrackId == original.TrackId && !string.Equals(a: previous.Name, b: reviewed.Name, comparisonType: StringComparison.Ordinal))
                        status = $"Track '{reviewed.Name}' saved as a new track. Track '{previous.Name}' retained.";
            try
            {
                await RefreshScanAsync();
                foreach (ScanTrack saved in _scanState!.Tracks)
                    if (string.Equals(a: saved.Name, b: reviewed.Name, comparisonType: StringComparison.Ordinal))
                    {
                        dgScanTracks.SelectedItem = saved;
                        FillScanEditor(track: saved);
                        break;
                    }
                txtScanStatus.Text = status;
            }
            catch (Exception ex)
            {
                txtScanStatus.Text = status + $" Refresh failed: {ex.Message}. Use Refresh.";
            }
        }, returnToTracks: true);
    }
    #endregion
}
