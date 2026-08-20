using System.Collections.Generic;

namespace GNA_DLRreport
{
    #region Track Geometry Import Configuration

    public sealed class TrackGeometryImportConfiguration
    {
        #region Global Rail Markers

        public List<string> StartMarkers { get; init; } =
            new();

        public List<string> EndMarkers { get; init; } =
            new();

        #endregion


        #region Worksheet Profiles

        public List<TrackGeometryWorksheetProfile> WorksheetProfiles { get; init; } =
            new();

        #endregion
    }


    public sealed class TrackGeometryWorksheetProfile
    {
        #region Worksheet Definition

        public string WorksheetName { get; init; } =
            string.Empty;

        public int StartRow { get; init; }

        public string PrimaryRailColumn { get; init; } =
            string.Empty;

        public string SecondaryRailColumn { get; init; } =
            string.Empty;

        public string TrackLabelColumn { get; init; } =
            string.Empty;

        #endregion
    }

    #endregion
}
