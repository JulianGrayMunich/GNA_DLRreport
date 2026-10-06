#region Sequential Debug Histogram Flyout
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
namespace GNA_DLRreport;

public sealed class ScanHeightHistogramWindow : Window
{
    public ScanHeightHistogramWindow(ScanHeightHistogram histogram)
    {
        ArgumentNullException.ThrowIfNull(argument: histogram);
        Title = (histogram.IsFiltered ? "Filtered subset — " : "Unfiltered scan elevations — ") + histogram.PointId;
        Width = 1050; Height = 620; MinWidth = 650; MinHeight = 550;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        DockPanel layout = new() { Margin = new Thickness(uniformLength: 14) };
        Content = layout;
        TextBlock title = new() { Text = $"{histogram.PointId} — {histogram.Count:N0} points — {(histogram.IsFiltered ? "final retained subset" : "before all height filters")}.\n{histogram.BandWidthMetres * 1000:0} mm bands [lower, upper). Elevation increases left to right; vertical axis is point count. Hover over a bar for its range and count.\nMean elevation: {(histogram.Mean.HasValue ? histogram.Mean.Value.ToString(format: "F6", provider: CultureInfo.InvariantCulture) + " m" : "not available (empty set)")}", TextWrapping = TextWrapping.Wrap, Margin = new Thickness(left: 0, top: 0, right: 0, bottom: 12) };
        DockPanel.SetDock(element: title, dock: Dock.Top); layout.Children.Add(element: title);
        StackPanel buttons = new() { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(left: 0, top: 12, right: 0, bottom: 0) };
        Button next = new() { Content = "Next", Width = 100, Height = 30, IsDefault = true };
        next.Click += (_, _) => DialogResult = true;
        Button abort = new() { Content = "Abort", Width = 100, Height = 30, IsCancel = true, Margin = new Thickness(left: 12, top: 0, right: 0, bottom: 0) };
        abort.Click += (_, _) => DialogResult = false;
        buttons.Children.Add(element: next); buttons.Children.Add(element: abort);
        DockPanel.SetDock(element: buttons, dock: Dock.Bottom); layout.Children.Add(element: buttons);
        if (histogram.Count == 0) { layout.Children.Add(element: new TextBlock { Text = "No scan points in this set.", FontSize = 20 }); return; }
        long first = histogram.Bands.Keys.Min(), last = histogram.Bands.Keys.Max(), maximum = histogram.Bands.Values.Max();
        double span = (double)last - first + 1;
        double width = Math.Clamp(value: span * 8d, min: 900d, max: 30000d);
        Canvas chart = new() { Width = width + 90, Height = 410, Background = Brushes.White };
        ScrollViewer scroll = new() { Content = chart, HorizontalScrollBarVisibility = ScrollBarVisibility.Auto, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };
        layout.Children.Add(element: scroll);
        const double left = 70, top = 15, plotHeight = 320;
        void Label(string text, double x, double y) { TextBlock label = new() { Text = text, FontSize = 12 }; Canvas.SetLeft(element: label, length: x); Canvas.SetTop(element: label, length: y); chart.Children.Add(element: label); }
        for (int i = 0; i <= 4; i++)
        {
            double y = top + plotHeight * (1 - i / 4d);
            chart.Children.Add(element: new Line { X1 = left, X2 = left + width, Y1 = y, Y2 = y, Stroke = Brushes.LightGray });
            Label(text: (maximum * i / 4d).ToString(format: "0.##", provider: CultureInfo.InvariantCulture), x: 0, y: y - 8);
        }
        foreach (var band in histogram.Bands)
        {
            double barHeight = plotHeight * band.Value / maximum;
            Rectangle bar = new() { Width = Math.Max(val1: .5d, val2: width / span - .5d), Height = barHeight, Fill = Brushes.SteelBlue,
                ToolTip = $"{band.Key * histogram.BandWidthMetres:F3} to {(band.Key + 1) * histogram.BandWidthMetres:F3} m (upper excluded): {band.Value:N0} points" };
            Canvas.SetLeft(element: bar, length: left + ((double)band.Key - first) * width / span);
            Canvas.SetTop(element: bar, length: top + plotHeight - barHeight); chart.Children.Add(element: bar);
        }
        if (histogram.Mean.HasValue)
        {
            double meanX = left + (histogram.Mean.Value / histogram.BandWidthMetres - first) * width / span;
            chart.Children.Add(element: new Line { X1 = meanX, X2 = meanX, Y1 = top, Y2 = top + plotHeight,
                Stroke = Brushes.DarkRed, StrokeThickness = 2, StrokeDashArray = new DoubleCollection { 4, 2 }, ToolTip = $"Mean: {histogram.Mean.Value:F6} m" });
        }
        for (int i = 0; i <= 5; i++) Label(text: ((first + span * i / 5d) * histogram.BandWidthMetres).ToString(format: "F3", provider: CultureInfo.InvariantCulture), x: left + width * i / 5d - 20, y: 345);
        Label(text: "Elevation (m)", x: left + width / 2 - 35, y: 375);
    }
}
#endregion

