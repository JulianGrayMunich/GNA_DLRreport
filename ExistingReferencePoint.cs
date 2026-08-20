using System;
using System.Collections.Generic;
using System.Text;

namespace GNA_DLRreport
{
    #region Existing Reference Point

    public sealed class ExistingReferencePoint
    {
        #region Database Identity

        public int PointName_ID { get; init; }

        public int Project_ID { get; init; }

        #endregion


        #region Point Identification

        public string PointName { get; init; } =
            string.Empty;

        public string ReplacementName { get; init; } =
            string.Empty;

        #endregion


        #region Reference Coordinates

        public decimal Easting { get; init; }

        public decimal Northing { get; init; }

        public decimal Height { get; init; }

        #endregion
    }

    #endregion
}
