#region Supported Height Bands
namespace GNA_DLRreport;
public static partial class TrackScan
{
    public const double MaximumReferenceDeviationMetres = 1.000d;
    private static void ValidateRejectionCount(int rejectionCount)
    {
        if (rejectionCount is not (5 or 10 or 15 or 20))
            throw new ArgumentOutOfRangeException(paramName: nameof(rejectionCount), message: "Select a rejection count of 5, 10, 15 or 20.");
    }
    private sealed class SurfaceRange
    {
        public int RejectionCount { get; }
        public SurfaceRange(int rejectionCount) { RejectionCount = rejectionCount; }
        private const double BandWidthMetres = 0.001d;
        private readonly Dictionary<long, (long Count, double Maximum)> _bands = new();
        public long Count { get; private set; }
        public long ReferenceRejectedCount { get; private set; }
        public double? Maximum { get; private set; }
        private static long Band(double height) => checked((long)Math.Floor(d: height / BandWidthMetres + 1e-8d));
        public void Observe(double height, double referenceHeight)
        {
            Count++;
            if (Math.Abs(value: height - referenceHeight) > MaximumReferenceDeviationMetres + BoundaryArithmeticAllowance)
            { ReferenceRejectedCount++; return; }
            long key = Band(height: height);
            if (_bands.TryGetValue(key: key, value: out var band))
                _bands[key] = (checked(band.Count + 1), Math.Max(val1: band.Maximum, val2: height));
            else _bands.Add(key: key, value: (1, height));
        }
        public void SelectSupportedMaximum()
        {
            Maximum = null;
            foreach (var band in _bands.Values)
                if (band.Count >= RejectionCount)
                    Maximum = Maximum.HasValue ? Math.Max(val1: Maximum.Value, val2: band.Maximum) : band.Maximum;
        }
        public bool IsSupported(double height) => _bands.TryGetValue(key: Band(height: height), value: out var band) && band.Count >= RejectionCount;
    }
}
#endregion
