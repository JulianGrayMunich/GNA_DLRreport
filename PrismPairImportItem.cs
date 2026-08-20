using System;
using System.Collections.Generic;
using System.Text;


namespace GNA_DLRreport
{
    #region Prism Pair Import Item

    public sealed class PrismPairImportItem
    {
        #region Worksheet Source

        public int SourceRow { get; init; }

        public int RailSection { get; init; }

        public string TrackName { get; init; } =
            string.Empty;

        public int PairOrder { get; init; }

        public string SourceColumn1Value { get; init; } =
            string.Empty;

        public string SourceColumn2Value { get; init; } =
            string.Empty;

        public string SourceColumn3Value { get; init; } =
            string.Empty;

        #endregion


        #region Resolved Prism Pair

        public string LeftPointName { get; init; } =
            string.Empty;

        public string RightPointName { get; init; } =
            string.Empty;

        #endregion


        #region Validation

        public bool IsValid { get; set; }

        public string ValidationStatus { get; set; } =
            string.Empty;

        #endregion
    }

    #endregion
}


