using System;
using System.Collections.Generic;
using System.Text;

namespace GNA_DLRreport
{
    #region Reference Coordinate Import Item

    public sealed class ReferenceCoordinateImportItem
    {
        #region Source Information

        public int SourceRow { get; init; }

        public string RawLine { get; init; } =
            string.Empty;

        #endregion


        #region Point Identification

        public string PointName { get; set; } =
            string.Empty;

        public string ReplacementName { get; set; } =
            string.Empty;

        #endregion


        #region Coordinate Display Values

        // Raw/normalised text is retained for the preview so malformed
        // numeric values can still be displayed to the user.

        public string EastingText { get; set; } =
            string.Empty;

        public string NorthingText { get; set; } =
            string.Empty;

        public string HeightText { get; set; } =
            string.Empty;

        #endregion


        #region Parsed Coordinate Values

        public decimal? Easting { get; set; }

        public decimal? Northing { get; set; }

        public decimal? Height { get; set; }

        #endregion


        #region Validation

        public bool IsValid { get; set; }

        public string ValidationStatus { get; set; } =
            string.Empty;

        #endregion
    }

    #endregion
}
