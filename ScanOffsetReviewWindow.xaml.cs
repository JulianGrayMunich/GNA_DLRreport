#region Stored Offset Review
using System.Windows;
namespace GNA_DLRreport;
public partial class ScanOffsetReviewWindow : Window
{
    public ScanOffsetReviewWindow(ScanProcessingResult result)
    {
        InitializeComponent();
        Language = System.Windows.Markup.XmlLanguage.GetLanguage(ietfLanguageTag: "en-GB");
        gridOffsets.ItemsSource = result.Offsets;
        int saved = result.Offsets.Count(predicate: offset => offset.ScanPointCount > 0);
        txtSummary.Text = $"{saved:N0} offsets saved; {result.Offsets.Count - saved:N0} polygons had no accepted points. " +
            "ScanOffset = reference ToR height − filtered scan mean. Existing offsets are retained where no new value can be computed.";
    }
}
#endregion
