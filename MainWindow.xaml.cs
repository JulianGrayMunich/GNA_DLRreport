#region System Preparation

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Line = System.Windows.Shapes.Line;
using Rectangle = System.Windows.Shapes.Rectangle;
using IOPath = System.IO.Path;
using GNAgeneraltools;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using OfficeOpenXml;

#endregion

namespace GNA_DLRreport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml.
    /// </summary>
    public partial class MainWindow : Window
    {

        #region Application Footer

        public string CopyrightText
        {
            get
            {
                DateTime buildDate =
                    GetBuildDate();

                int year =
                    buildDate.Year;

                return
                    $"© {year} GNA Software — Built {buildDate:yyyy-MM-dd}";
            }
        }


        private static DateTime GetBuildDate()
        {
            #region Resolve Executable Build Date

            string assemblyLocation =
                System.Reflection.Assembly
                    .GetExecutingAssembly()
                    .Location;

            if (string.IsNullOrWhiteSpace(
                value: assemblyLocation))
            {
                return DateTime.Today;
            }

            DateTime buildDate =
                File.GetLastWriteTime(
                    path: assemblyLocation);

            return buildDate.Date;

            #endregion
        }

        #endregion


        #region Setting State
        private readonly gnaTools gnaT =
            new();

        #endregion


        #region Registry Configuration

        private const string RegistryPath =
            @"SOFTWARE\GNASoftware\DLRReport";

        private const string RegistryDatabaseConnectionString =
            "DatabaseConnectionString";

        private const string RegistryActiveProjectId =
            "ActiveProjectID";

        private const string RegistryActiveProjectName =
            "ActiveProjectName";

        #endregion


        #region Database Configuration

        private const string TrackGeometryDatabaseName =
            "DBTrackGeometry";

        // Operational database connection. This is updated only after a
        // successful connection test or from the persisted startup default.
        private string _validatedDatabaseConnectionString =
            string.Empty;

        #endregion


        #region Project Configuration State

        // Project records currently displayed in the Project Management DataGrid.
        private readonly ObservableCollection<ProjectConfigurationItem> _projectItems =
            new();


        // Project start dates are stored separately from ProjectConfigurationItem
        // so this replacement file does not require changes to the existing model
        // class used elsewhere in the solution.
        private readonly Dictionary<int, DateTime> _projectStartDates =
            new();

        private DateTime _pendingProjectStartDate =
            DateTime.Today;

        private bool _isUpdatingProjectStartDateControl;


        private int? _projectStartCalendarTargetProjectId;


        // ---------------------------------------------------------------------
        // MULTI-INSTANCE RUNTIME STATE RULE
        //
        // Registry values are persisted startup defaults only.
        //
        // Each running application instance owns its own operational project
        // context through:
        //
        //     _activeProjectId
        //     _activeProjectName
        //
        // Registry values must not be repeatedly read during normal processing.
        // ---------------------------------------------------------------------

        private int? _activeProjectId;

        private string _activeProjectName =
            string.Empty;


        // ---------------------------------------------------------------------
        // CROSS-INSTANCE ACTIVE PROJECT LOCK
        //
        // Each running instance holding an active project maintains one dedicated
        // SQL connection.
        //
        // A Shared SQL application lock is held on:
        //
        //     GNA_DLRreport:Project:<Project_ID>
        //
        // Multiple running instances may therefore use the same active project.
        //
        // Project deletion requires an Exclusive lock on the same
        // resource. The Exclusive lock cannot be obtained while any running
        // instance holds a Shared lock.
        //
        // The dedicated lock connection uses Pooling = false so disposing the
        // connection terminates its SQL session and releases its Session-owned
        // application lock.
        // ---------------------------------------------------------------------

        private SqlConnection? _activeProjectLockConnection;

        private const string ProjectLockResourcePrefix =
            "GNA_DLRreport:Project:";

        #endregion


        #region Chart Configuration State

        private const int DefaultChartRecentDays =
            14;

        // Fixed report colours. These are intentionally not exposed in the UI.
        private const string ChartGreenBandColourHex =
            "#DFF2D8";

        private const string ChartAmberBandColourHex =
            "#FFF2CC";

        private const string ChartRedBandColourHex =
            "#F4CCCC";

        private const string ChartAboveRedBandColourHex =
            "#E7E6E6";

        private const string ChartZeroAxisColourHex =
            "#000000";

        private readonly ObservableCollection<ChartTypeUiItem> _chartTypes =
            new();

        private readonly ObservableCollection<ChartSeriesUiItem> _chartSeries =
            new();


        private readonly ObservableCollection<ExistingChartUiItem> _existingCharts =
            new();


        private readonly ObservableCollection<ChartTemplateUiItem> _chartTemplates =
            new();

        private const int DefaultChartWidthMm = 150;

        private const int DefaultChartHeightMm = 80;

        private const int DefaultChartResolutionDpi = 300;

        private const string DefaultChartFontFamily = "Arial";

        private const string ChartStartDateModeReportStart = "ReportStart";

        private const string ChartStartDateModeProjectStart = "ProjectStart";

        private IChartPreviewRenderer _chartPreviewRenderer = null!;

        private interface IChartPreviewRenderer
        {
            void Render();
        }

        private sealed class WpfCanvasChartPreviewRenderer : IChartPreviewRenderer
        {
            private readonly MainWindow _owner;

            public WpfCanvasChartPreviewRenderer(
                MainWindow owner)
            {
                _owner = owner ?? throw new ArgumentNullException(
                    paramName: nameof(owner));
            }

            public void Render()
            {
                _owner.RenderBlankChartCanvas();
            }
        }

        private int? _loadedChartTemplateId;

        private sealed class ChartTemplateUiItem
        {
            public int ChartTemplateId { get; init; }

            public string TemplateName { get; init; } =
                string.Empty;

            public string ChartTypeKey { get; init; } =
                string.Empty;

            public string DisplayText =>
                TemplateName;
        }

        private int? _loadedChartDefinitionId;

        private int? _loadedChartNumber;

        private sealed class ExistingChartUiItem
        {
            public int ChartDefinitionId { get; init; }

            public int ChartNumber { get; init; }

            public string ChartName { get; init; } =
                string.Empty;

            public string ChartTypeKey { get; init; } =
                string.Empty;

            public string DisplayText =>
                $"Chart_{ChartNumber:0000} - {ChartName}";
        }

        private sealed class ChartTypeUiItem
        {
            public string Key { get; init; } =
                string.Empty;

            public string DisplayName { get; init; } =
                string.Empty;

            public string SourceTable { get; init; } =
                string.Empty;

            public string EntityKind { get; init; } =
                string.Empty;

            public string VerticalAxisTitle { get; init; } =
                string.Empty;

            public string Unit { get; init; } =
                string.Empty;

            public bool IsFullyDefined { get; init; } =
                true;

            public IReadOnlyList<string> DataElements { get; init; } =
                Array.Empty<string>();

            public override string ToString()
            {
                return DisplayName;
            }
        }

        private sealed class ChartSeriesUiItem
        {
            public int EntityId { get; set; }

            public int DisplayOrder { get; set; }

            public string EntityDisplayName { get; set; } =
                string.Empty;

            public string DataElementDisplayName { get; set; } =
                string.Empty;

            public string LegendText { get; set; } =
                string.Empty;

            public string ColourHex { get; set; } =
                string.Empty;

            public double LineWidth { get; set; } =
                2.0;

            public double MarkerSize { get; set; } =
                4.0;
        }

        private sealed class ChartEntityUiItem
        {
            public int EntityId { get; init; }

            public string DisplayName { get; init; } =
                string.Empty;

            public override string ToString()
            {
                return DisplayName;
            }
        }

        private sealed record ChartPreviewPoint(
            DateTime UtcTime,
            double Value);

        private sealed class ChartPreviewSeries
        {
            public string LegendText { get; init; } =
                string.Empty;

            public Brush Stroke { get; init; } =
                Brushes.Blue;

            public double LineWidth { get; init; } =
                2.0;

            public List<ChartPreviewPoint> Points { get; } =
                new();
        }

        private readonly List<ChartPreviewSeries> _chartPreviewSeries =
            new();

        private DateTime _chartPreviewStartUtc;

        private DateTime _chartPreviewEndUtcExclusive;

        #endregion


        #region Reference Coordinate Import State

        private readonly ObservableCollection<ReferenceCoordinateImportItem>
            _referenceImportItems =
                new();

        private string _selectedReferenceCsvPath =
            string.Empty;


        // ---------------------------------------------------------------------
        // IMPORT PROJECT CONTEXT
        //
        // The Project_ID is captured when the CSV is selected.
        //
        // The complete validation operation is then tied to that Project_ID.
        // Changing the active project clears the current import.
        //
        // Registry values are never used as operational import context.
        // ---------------------------------------------------------------------

        private int? _referenceImportProjectId;

        private string _referenceImportProjectName =
            string.Empty;


        // ---------------------------------------------------------------------
        // DUPLICATE COORDINATE RULE
        //
        // Duplicate proximity is based on horizontal E/N separation only.
        //
        // Duplicate when:
        //      distance < 0.05 m
        //
        // Exactly 0.0500 m is permitted.
        // ---------------------------------------------------------------------

        private const decimal ReferenceDuplicateCoordinateTolerance =
            0.05m;

        private const decimal ReferenceDuplicateCoordinateToleranceSquared =
            ReferenceDuplicateCoordinateTolerance *
            ReferenceDuplicateCoordinateTolerance;

        #endregion


        #region Geotechnical Sensor Import State

        private readonly ObservableCollection<GeotechSensorImportItem>
            _geotechImportItems =
                new();

        private string _selectedGeotechCsvPath =
            string.Empty;

        private int? _geotechImportProjectId;

        private string _geotechImportProjectName =
            string.Empty;

        #endregion


        #region Prism Pair Import State

        private const string PrismPairImportProfileFileName =
            "TrackGeometryImportProfiles.json";

        private static readonly JsonSerializerOptions PrismPairJsonSerializerOptions =
            new()
            {
                PropertyNameCaseInsensitive =
                    true
            };


        private const string PrismPairRoleRightRail =
            "Right (Primary) Rail";

        private const string PrismPairRoleLeftRail =
            "Left Rail";

        private const string PrismPairRoleRailTag =
            "Rail Tag";


        private const string PrismPairImportLockResourcePrefix =
            "GNA_DLRreport:PrismPairImport:";


        private readonly List<PrismPairImportItem> _prismPairImportItems =
            new();


        private string _selectedPrismPairWorkbookPath =
            string.Empty;

        private TrackGeometryImportConfiguration? _prismPairImportConfiguration;

        private TrackGeometryWorksheetProfile? _selectedPrismPairWorksheetProfile;

        private bool _isUpdatingPrismPairColumnRoles;


        private int? _prismPairImportProjectId;

        private string _prismPairImportProjectName =
            string.Empty;

        #endregion


        #region Prism Array Configuration State

        private const int PrismArrayTypePrismCrackGauge =
            3;

        private const int PrismArrayTypePrismTilt =
            4;

        private const string PrismArrayDefinitionLockResourcePrefix =
            "GNA_DLRreport:PrismArrayDefinition:";

        private static readonly char[] PrismArrayPointRoles =
        {
            'A',
            'B',
            'C',
            'D',
            'E'
        };

        private static readonly char[] PrismCrackGaugePointRoles =
        {
            'A',
            'B'
        };

        private static readonly (char StartRole, char EndRole, string ChordName)[]
            TunnelConvergenceChordDefinitions =
        {
            ('A', 'B', "AB"),
            ('A', 'C', "AC"),
            ('A', 'D', "AD"),
            ('A', 'E', "AE"),
            ('B', 'C', "BC"),
            ('B', 'D', "BD"),
            ('B', 'E', "BE"),
            ('C', 'D', "CD"),
            ('C', 'E', "CE"),
            ('D', 'E', "DE")
        };

        private readonly ObservableCollection<PrismArrayTypeItem>
            _prismArrayTypes =
                new();

        private readonly ObservableCollection<PrismArrayAvailablePoint>
            _availablePrismArrayPoints =
                new();

        private readonly ObservableCollection<PrismArraySummaryItem>
            _committedPrismArrays =
                new();

        private readonly Dictionary<char, PrismArrayAvailablePoint>
            _prismArrayAssignments =
                new();

        private int? _prismArrayProjectId;

        private string _prismArrayProjectName =
            string.Empty;

        private int? _selectedPrismArrayType;

        private int? _editingPrismArrayId;

        private string _pendingPrismArrayName =
            string.Empty;

        #endregion



        #region Prism Array Configuration

        #region Array State Models

        private sealed class PrismArrayTypeItem
        {
            public int ArrayType_ID { get; init; }

            public string ArrayTypeName { get; init; } =
                string.Empty;
        }


        private sealed class PrismArrayAvailablePoint
        {
            public int PointName_ID { get; init; }

            public string PointName { get; init; } =
                string.Empty;

            public string ReplacementName { get; init; } =
                string.Empty;
        }


        private sealed class PrismArraySummaryItem
        {
            public int Array_ID { get; init; }

            public int ArrayType { get; init; }

            public string ArrayTypeName { get; init; } =
                string.Empty;

            public string ArrayName { get; init; } =
                string.Empty;

            public bool IsDeleted { get; set; }

            public bool IsIncluded
            {
                get => !IsDeleted;
                set => IsDeleted = !value;
            }
        }


        private sealed class PrismArrayDetailsItem
        {
            public int Array_ID { get; init; }

            public int Project_ID { get; init; }

            public int ArrayType { get; init; }

            public string ArrayTypeName { get; init; } =
                string.Empty;

            public string ArrayName { get; init; } =
                string.Empty;

            public bool IsDeleted { get; init; }

            public Dictionary<char, PrismArrayAvailablePoint> Points { get; } =
                new();
        }

        #endregion


        #region UTC Timestamp Standard

        private static DateTime RoundUtcToNearestSecond(
            DateTime utcTime)
        {
            #region Validate UTC Timestamp

            if (utcTime.Kind != DateTimeKind.Utc)
            {
                throw new ArgumentException(
                    message: "UTC timestamp required.",
                    paramName: nameof(utcTime));
            }

            #endregion


            #region Round To Nearest Whole Second

            long ticksPerSecond =
                TimeSpan.TicksPerSecond;

            long remainder =
                utcTime.Ticks % ticksPerSecond;

            long roundedTicks =
                utcTime.Ticks - remainder;

            if (remainder >= ticksPerSecond / 2)
            {
                roundedTicks +=
                    ticksPerSecond;
            }

            if (roundedTicks > DateTime.MaxValue.Ticks)
            {
                roundedTicks =
                    DateTime.MaxValue.Ticks -
                    (DateTime.MaxValue.Ticks % ticksPerSecond);
            }

            return new DateTime(
                ticks: roundedTicks,
                kind: DateTimeKind.Utc);

            #endregion
        }

        #endregion


        #region Array Type And Role Helpers

        private PrismArrayTypeItem GetPrismArrayType(
            int arrayTypeId)
        {
            #region Resolve Array Type From Database Lookup

            foreach (PrismArrayTypeItem arrayType
                in _prismArrayTypes)
            {
                if (arrayType.ArrayType_ID == arrayTypeId)
                {
                    return arrayType;
                }
            }

            throw new InvalidOperationException(
                $"Array type {arrayTypeId} is not defined.");

            #endregion
        }


        private void ValidatePrismArrayType(
            int arrayTypeId)
        {
            #region Validate Array Type Against Database Lookup

            _ =
                GetPrismArrayType(
                    arrayTypeId: arrayTypeId);

            #endregion
        }


        private static IReadOnlyList<char> GetRequiredPrismArrayPointRoles(
            int arrayType)
        {
            #region Resolve Required Point Roles

            return
                arrayType == PrismArrayTypePrismCrackGauge ||
                arrayType == PrismArrayTypePrismTilt
                    ? PrismCrackGaugePointRoles
                    : PrismArrayPointRoles;

            #endregion
        }


        private static bool PrismArrayTypeUsesPointRole(
            int arrayType,
            char pointRole)
        {
            #region Resolve Role Availability

            char validatedPointRole =
                ValidatePrismArrayPointRole(
                    pointRole: pointRole);

            foreach (char requiredPointRole
                in GetRequiredPrismArrayPointRoles(
                    arrayType: arrayType))
            {
                if (requiredPointRole == validatedPointRole)
                {
                    return true;
                }
            }

            return false;

            #endregion
        }


        private static char ValidatePrismArrayPointRole(
            char pointRole)
        {
            #region Validate Point Role

            char validatedPointRole =
                char.ToUpperInvariant(
                    c: pointRole);

            foreach (char allowedRole
                in PrismArrayPointRoles)
            {
                if (validatedPointRole == allowedRole)
                {
                    return validatedPointRole;
                }
            }

            throw new ArgumentOutOfRangeException(
                paramName: nameof(pointRole),
                message: "Array role must be A, B, C, D or E.");

            #endregion
        }

        #endregion


        #region Array Runtime State

        private void ResetPrismArrayConfigurationState()
        {
            #region Clear Array Project Context

            _prismArrayProjectId =
                null;

            _prismArrayProjectName =
                string.Empty;

            #endregion


            #region Clear Pending Array Definition

            _selectedPrismArrayType =
                null;

            _editingPrismArrayId =
                null;

            _pendingPrismArrayName =
                string.Empty;

            _prismArrayAssignments.Clear();

            #endregion


            #region Clear Array Lists

            _prismArrayTypes.Clear();

            _availablePrismArrayPoints.Clear();

            _committedPrismArrays.Clear();

            #endregion


            #region Clear Array User Interface

            if (IsInitialized)
            {
                cmbPrismArrayType.SelectedIndex =
                    -1;

                ClearPrismArrayDefinitionControls();

                dgCommittedPrismArrays.SelectedItem =
                    null;

                bool arraysTabActive =
                    tabPrismArrays.IsSelected;

                btnViewPrismArray.Visibility =
                    arraysTabActive
                        ? Visibility.Visible
                        : Visibility.Collapsed;

                btnViewPrismArray.IsEnabled =
                    arraysTabActive;
            }

            #endregion
        }


        private void ResetPendingPrismArrayDefinition()
        {
            #region Preserve Pending Points For Return

            List<PrismArrayAvailablePoint> pointsToReturn =
                new();

            foreach (PrismArrayAvailablePoint assignedPoint
                in _prismArrayAssignments.Values)
            {
                pointsToReturn.Add(
                    item: assignedPoint);
            }

            #endregion


            #region Clear Pending Definition

            _pendingPrismArrayName =
                string.Empty;

            _editingPrismArrayId =
                null;

            _prismArrayAssignments.Clear();

            #endregion


            #region Return Points To Available List

            foreach (PrismArrayAvailablePoint point
                in pointsToReturn)
            {
                InsertAvailablePrismArrayPointSorted(
                    point: point);
            }

            #endregion
        }


        private void SetPendingPrismArrayDefinition(
            int arrayType,
            string arrayName)
        {
            #region Validate Definition Header

            ValidatePrismArrayType(
                arrayTypeId: arrayType);

            string validatedArrayName =
                arrayName?.Trim()
                ?? string.Empty;

            if (string.IsNullOrWhiteSpace(validatedArrayName))
            {
                throw new ArgumentException(
                    message: "Array name required.",
                    paramName: nameof(arrayName));
            }

            if (validatedArrayName.Length > 200)
            {
                throw new ArgumentException(
                    message: "Array name exceeds 200 characters.",
                    paramName: nameof(arrayName));
            }

            #endregion


            #region Store Definition Header

            _selectedPrismArrayType =
                arrayType;

            _pendingPrismArrayName =
                validatedArrayName;

            #endregion
        }

        #endregion


        #region Array Type Lookup

        private async Task LoadPrismArrayTypesAsync()
        {
            #region Preserve Current Array Type Selection

            int? preservedArrayTypeId =
                cmbPrismArrayType.SelectedItem
                    is PrismArrayTypeItem selectedArrayType
                        ? selectedArrayType.ArrayType_ID
                        : _selectedPrismArrayType;

            #endregion


            #region Define Array Type Query

            const string arrayTypeSql = """
                SELECT
                    [ArrayType_ID],
                    [ArrayTypeName]
                FROM [dbo].[PrismArrayType]
                ORDER BY
                    [ArrayType_ID];
                """;

            #endregion


            #region Load Array Types

            List<PrismArrayTypeItem> arrayTypes =
                new();

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand arrayTypeCommand =
                new(
                    cmdText: arrayTypeSql,
                    connection: databaseConnection);

            await using SqlDataReader reader =
                await arrayTypeCommand.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                arrayTypes.Add(
                    item:
                        new PrismArrayTypeItem
                        {
                            ArrayType_ID =
                                reader.GetInt32(
                                    i: 0),

                            ArrayTypeName =
                                reader.GetString(
                                    i: 1)
                        });
            }

            #endregion


            #region Validate Array Type Lookup

            if (arrayTypes.Count == 0)
            {
                throw new InvalidOperationException(
                    "No array types are defined.");
            }

            #endregion


            #region Refresh Array Type User Interface

            _prismArrayTypes.Clear();

            foreach (PrismArrayTypeItem arrayType
                in arrayTypes)
            {
                _prismArrayTypes.Add(
                    item: arrayType);
            }

            cmbPrismArrayType.ItemsSource =
                _prismArrayTypes;

            if (preservedArrayTypeId.HasValue)
            {
                cmbPrismArrayType.SelectedValue =
                    preservedArrayTypeId.Value;
            }
            else
            {
                cmbPrismArrayType.SelectedIndex =
                    -1;
            }

            #endregion
        }

        #endregion


        #region Available Array Points

        private async Task InitialisePrismArrayConfigurationForActiveProjectAsync()
        {
            #region Validate Active Project

            if (!_activeProjectId.HasValue ||
                string.IsNullOrWhiteSpace(_activeProjectName))
            {
                ResetPrismArrayConfigurationState();

                throw new InvalidOperationException(
                    "Select an active project.");
            }

            #endregion


            #region Establish Array Project Context

            ResetPrismArrayConfigurationState();

            _prismArrayProjectId =
                _activeProjectId.Value;

            _prismArrayProjectName =
                _activeProjectName;

            #endregion


            #region Load Array Types

            await LoadPrismArrayTypesAsync();

            #endregion


            #region Load Array Lists

            await ReloadPrismArrayConfigurationListsAsync();

            #endregion
        }


        private async Task ReloadPrismArrayConfigurationListsAsync()
        {
            #region Validate Array Project Context

            if (!_prismArrayProjectId.HasValue ||
                string.IsNullOrWhiteSpace(_prismArrayProjectName))
            {
                throw new InvalidOperationException(
                    "Array project unavailable.");
            }

            if (!_activeProjectId.HasValue ||
                _activeProjectId.Value != _prismArrayProjectId.Value)
            {
                throw new InvalidOperationException(
                    "Active project changed.");
            }

            int projectId =
                _prismArrayProjectId.Value;

            #endregion


            #region Open Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            #endregion


            #region Load Available Points

            List<PrismArrayAvailablePoint> availablePoints =
                await LoadAvailablePrismArrayPointsAsync(
                    projectId: projectId,
                    databaseConnection: databaseConnection,
                    transaction: null);

            HashSet<int> assignedPointIds =
                new();

            foreach (PrismArrayAvailablePoint assignedPoint
                in _prismArrayAssignments.Values)
            {
                assignedPointIds.Add(
                    item: assignedPoint.PointName_ID);
            }

            _availablePrismArrayPoints.Clear();

            foreach (PrismArrayAvailablePoint availablePoint
                in availablePoints)
            {
                if (!assignedPointIds.Contains(
                    item: availablePoint.PointName_ID))
                {
                    _availablePrismArrayPoints.Add(
                        item: availablePoint);
                }
            }

            #endregion


            #region Load Committed Arrays

            List<PrismArraySummaryItem> committedArrays =
                await LoadCommittedPrismArraysAsync(
                    projectId: projectId,
                    databaseConnection: databaseConnection,
                    transaction: null);

            _committedPrismArrays.Clear();

            foreach (PrismArraySummaryItem committedArray
                in committedArrays)
            {
                _committedPrismArrays.Add(
                    item: committedArray);
            }

            #endregion
        }


        private static async Task<List<PrismArrayAvailablePoint>>
            LoadAvailablePrismArrayPointsAsync(
                int projectId,
                SqlConnection databaseConnection,
                SqlTransaction? transaction)
        {
            #region Define Available Point Query

            const string availablePointSql = """
                SELECT
                    PN.[PointName_ID],
                    PN.[PointName],
                    PN.[ReplacementName]
                FROM [dbo].[PointName] AS PN
                INNER JOIN [dbo].[CoordinatesReference] AS CR
                    ON CR.[PointName_ID] = PN.[PointName_ID]
                    AND CR.[IsDeleted] = 0
                WHERE
                    PN.[Project_ID] = @Project_ID
                    AND PN.[IsDeleted] = 0
                ORDER BY
                    PN.[ReplacementName],
                    PN.[PointName];
                """;

            #endregion


            #region Execute Available Point Query

            List<PrismArrayAvailablePoint> availablePoints =
                new();

            await using SqlCommand availablePointCommand =
                transaction is null
                    ? new SqlCommand(
                        cmdText: availablePointSql,
                        connection: databaseConnection)
                    : new SqlCommand(
                        cmdText: availablePointSql,
                        connection: databaseConnection,
                        transaction: transaction);

            availablePointCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            await using SqlDataReader reader =
                await availablePointCommand.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                availablePoints.Add(
                    item:
                        new PrismArrayAvailablePoint
                        {
                            PointName_ID =
                                reader.GetInt32(
                                    i: 0),

                            PointName =
                                reader.GetString(
                                    i: 1),

                            ReplacementName =
                                reader.GetString(
                                    i: 2)
                        });
            }

            return availablePoints;

            #endregion
        }


        private void InsertAvailablePrismArrayPointSorted(
            PrismArrayAvailablePoint point)
        {
            #region Avoid Duplicate Available Point

            foreach (PrismArrayAvailablePoint availablePoint
                in _availablePrismArrayPoints)
            {
                if (availablePoint.PointName_ID == point.PointName_ID)
                {
                    return;
                }
            }

            #endregion


            #region Find Sorted Insert Position

            int insertIndex =
                _availablePrismArrayPoints.Count;

            for (int index = 0;
                 index < _availablePrismArrayPoints.Count;
                 index++)
            {
                PrismArrayAvailablePoint existingPoint =
                    _availablePrismArrayPoints[index];

                int replacementComparison =
                    string.Compare(
                        strA: point.ReplacementName,
                        strB: existingPoint.ReplacementName,
                        comparisonType: StringComparison.OrdinalIgnoreCase);

                if (replacementComparison < 0)
                {
                    insertIndex =
                        index;

                    break;
                }

                if (replacementComparison == 0 &&
                    string.Compare(
                        strA: point.PointName,
                        strB: existingPoint.PointName,
                        comparisonType: StringComparison.OrdinalIgnoreCase) < 0)
                {
                    insertIndex =
                        index;

                    break;
                }
            }

            #endregion


            #region Insert Available Point

            _availablePrismArrayPoints.Insert(
                index: insertIndex,
                item: point);

            #endregion
        }

        #endregion


        #region A-E Point Assignment

        private void AssignPrismArrayPoint(
            char pointRole,
            int pointNameId)
        {
            #region Validate Role And Point

            char validatedPointRole =
                ValidatePrismArrayPointRole(
                    pointRole: pointRole);

            PrismArrayAvailablePoint? selectedPoint =
                null;

            foreach (PrismArrayAvailablePoint availablePoint
                in _availablePrismArrayPoints)
            {
                if (availablePoint.PointName_ID == pointNameId)
                {
                    selectedPoint =
                        availablePoint;

                    break;
                }
            }

            if (selectedPoint is null)
            {
                throw new InvalidOperationException(
                    "Point unavailable.");
            }

            #endregion


            #region Replace Existing Role Assignment

            if (_prismArrayAssignments.TryGetValue(
                key: validatedPointRole,
                value: out PrismArrayAvailablePoint? existingAssignment))
            {
                _prismArrayAssignments.Remove(
                    key: validatedPointRole);

                InsertAvailablePrismArrayPointSorted(
                    point: existingAssignment);
            }

            #endregion


            #region Assign Selected Point

            _availablePrismArrayPoints.Remove(
                item: selectedPoint);

            _prismArrayAssignments.Add(
                key: validatedPointRole,
                value: selectedPoint);

            #endregion
        }


        private void ClearPrismArrayPointAssignment(
            char pointRole)
        {
            #region Validate Point Role

            char validatedPointRole =
                ValidatePrismArrayPointRole(
                    pointRole: pointRole);

            #endregion


            #region Clear Assignment

            if (!_prismArrayAssignments.TryGetValue(
                key: validatedPointRole,
                value: out PrismArrayAvailablePoint? assignedPoint))
            {
                return;
            }

            _prismArrayAssignments.Remove(
                key: validatedPointRole);

            InsertAvailablePrismArrayPointSorted(
                point: assignedPoint);

            #endregion
        }


        private PrismArrayAvailablePoint? GetPrismArrayPointAssignment(
            char pointRole)
        {
            #region Get Assignment

            char validatedPointRole =
                ValidatePrismArrayPointRole(
                    pointRole: pointRole);

            return _prismArrayAssignments.TryGetValue(
                key: validatedPointRole,
                value: out PrismArrayAvailablePoint? assignedPoint)
                    ? assignedPoint
                    : null;

            #endregion
        }

        #endregion


        #region Prism Array Commit

        private async Task<int> CommitCurrentPrismArrayDefinitionAsync()
        {
            #region Validate Runtime Context

            if (!_prismArrayProjectId.HasValue ||
                string.IsNullOrWhiteSpace(_prismArrayProjectName))
            {
                throw new InvalidOperationException(
                    "Array project unavailable.");
            }

            if (!_activeProjectId.HasValue ||
                _activeProjectId.Value != _prismArrayProjectId.Value)
            {
                throw new InvalidOperationException(
                    "Active project changed.");
            }

            if (!_selectedPrismArrayType.HasValue)
            {
                throw new InvalidOperationException(
                    "Select an array type.");
            }

            int arrayType =
                _selectedPrismArrayType.Value;

            ValidatePrismArrayType(
                arrayTypeId: arrayType);

            string arrayName =
                _pendingPrismArrayName.Trim();

            if (string.IsNullOrWhiteSpace(arrayName))
            {
                throw new InvalidOperationException(
                    "Array name required.");
            }

            if (arrayName.Length > 200)
            {
                throw new InvalidOperationException(
                    "Array name exceeds 200 characters.");
            }

            Dictionary<char, int> pointAssignments =
                CapturePrismArrayPointAssignments(
                    arrayType: arrayType);

            int projectId =
                _prismArrayProjectId.Value;

            string projectName =
                _prismArrayProjectName;

            #endregion


            #region Open Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            #endregion


            #region Begin Array Definition Transaction

            using SqlTransaction transaction =
                databaseConnection.BeginTransaction(
                    iso:
                        System.Data.IsolationLevel.Serializable);

            bool transactionCommitted =
                false;

            #endregion


            try
            {
                #region Acquire Array Definition Lock

                await AcquirePrismArrayDefinitionTransactionLockAsync(
                    projectId: projectId,
                    databaseConnection: databaseConnection,
                    transaction: transaction);

                #endregion


                #region Revalidate Target Project

                await ValidateImportTargetProjectAsync(
                    projectId: projectId,
                    expectedProjectName: projectName,
                    databaseConnection: databaseConnection,
                    transaction: transaction);

                #endregion


                #region Final Array Definition Validation

                await ValidatePrismArrayDefinitionForCommitAsync(
                    projectId: projectId,
                    existingArrayId: _editingPrismArrayId,
                    arrayName: arrayName,
                    arrayType: arrayType,
                    pointAssignments: pointAssignments,
                    databaseConnection: databaseConnection,
                    transaction: transaction);

                #endregion


                #region Insert Or Update Array Definition

                int arrayId;

                if (_editingPrismArrayId.HasValue)
                {
                    arrayId =
                        _editingPrismArrayId.Value;

                    await UpdatePrismArrayAsync(
                        arrayId: arrayId,
                        projectId: projectId,
                        arrayName: arrayName,
                        arrayType: arrayType,
                        databaseConnection: databaseConnection,
                        transaction: transaction);

                    await DeletePrismArrayPointsAsync(
                        arrayId: arrayId,
                        databaseConnection: databaseConnection,
                        transaction: transaction);
                }
                else
                {
                    arrayId =
                        await InsertPrismArrayAsync(
                            projectId: projectId,
                            arrayName: arrayName,
                            arrayType: arrayType,
                            databaseConnection: databaseConnection,
                            transaction: transaction);
                }

                #endregion


                #region Insert A-E Membership

                foreach (char pointRole
                    in GetRequiredPrismArrayPointRoles(
                        arrayType: arrayType))
                {
                    await InsertPrismArrayPointAsync(
                        arrayId: arrayId,
                        pointRole: pointRole,
                        pointNameId: pointAssignments[pointRole],
                        databaseConnection: databaseConnection,
                        transaction: transaction);
                }

                #endregion


                #region Commit Array Definition

                transaction.Commit();

                transactionCommitted =
                    true;

                #endregion


                #region Refresh Runtime State

                ResetPendingPrismArrayDefinition();

                await ReloadPrismArrayConfigurationListsAsync();

                return arrayId;

                #endregion
            }
            catch
            {
                #region Roll Back Array Definition

                if (!transactionCommitted)
                {
                    try
                    {
                        transaction.Rollback();
                    }
                    catch
                    {
                        // Preserve the original exception.
                    }
                }

                throw;

                #endregion
            }
        }


        private Dictionary<char, int> CapturePrismArrayPointAssignments(
            int arrayType)
        {
            #region Validate Required Point Assignment

            Dictionary<char, int> pointAssignments =
                new();

            HashSet<int> pointIds =
                new();

            foreach (char pointRole
                in GetRequiredPrismArrayPointRoles(
                    arrayType: arrayType))
            {
                if (!_prismArrayAssignments.TryGetValue(
                    key: pointRole,
                    value: out PrismArrayAvailablePoint? assignedPoint))
                {
                    throw new InvalidOperationException(
                        $"Point {pointRole} not assigned.");
                }

                if (!pointIds.Add(
                    item: assignedPoint.PointName_ID))
                {
                    throw new InvalidOperationException(
                        "Array points must be unique.");
                }

                pointAssignments.Add(
                    key: pointRole,
                    value: assignedPoint.PointName_ID);
            }

            return pointAssignments;

            #endregion
        }


        private static async Task AcquirePrismArrayDefinitionTransactionLockAsync(
            int projectId,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Array Definition Lock

            string lockResource =
                $"{PrismArrayDefinitionLockResourcePrefix}{projectId}";

            const string lockSql = """
                DECLARE @LockResult int;

                EXEC @LockResult = sys.sp_getapplock
                    @Resource = @Resource,
                    @LockMode = 'Exclusive',
                    @LockOwner = 'Transaction',
                    @LockTimeout = 0;

                SELECT @LockResult;
                """;

            #endregion


            #region Acquire Array Definition Lock

            await using SqlCommand lockCommand =
                new(
                    cmdText: lockSql,
                    connection: databaseConnection,
                    transaction: transaction);

            lockCommand.Parameters.Add(
                parameterName: "@Resource",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 255)
                .Value =
                    lockResource;

            object? lockResultValue =
                await lockCommand.ExecuteScalarAsync();

            if (lockResultValue is null ||
                lockResultValue == DBNull.Value)
            {
                throw new InvalidOperationException(
                    "Array lock unavailable.");
            }

            int lockResult =
                Convert.ToInt32(
                    value: lockResultValue,
                    provider: CultureInfo.InvariantCulture);

            if (lockResult < 0)
            {
                throw new InvalidOperationException(
                    "Another array definition is active.");
            }

            #endregion
        }


        private static async Task ValidatePrismArrayDefinitionForCommitAsync(
            int projectId,
            int? existingArrayId,
            string arrayName,
            int arrayType,
            Dictionary<char, int> pointAssignments,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Validate Array Name Uniqueness

            const string arrayNameSql = """
                SELECT COUNT(1)
                FROM [dbo].[PrismArray] WITH (UPDLOCK, HOLDLOCK)
                WHERE
                    [Project_ID] = @Project_ID
                    AND [ArrayName] = @ArrayName
                    AND [IsDeleted] = 0
                    AND
                    (
                        @ExistingArray_ID IS NULL
                        OR [Array_ID] <> @ExistingArray_ID
                    );
                """;

            await using (SqlCommand arrayNameCommand =
                new(
                    cmdText: arrayNameSql,
                    connection: databaseConnection,
                    transaction: transaction))
            {
                arrayNameCommand.Parameters.Add(
                    parameterName: "@Project_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value =
                        projectId;

                arrayNameCommand.Parameters.Add(
                    parameterName: "@ArrayName",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 200)
                    .Value =
                        arrayName;

                arrayNameCommand.Parameters.Add(
                    parameterName: "@ExistingArray_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value =
                        existingArrayId.HasValue
                            ? existingArrayId.Value
                            : DBNull.Value;

                int existingArrayCount =
                    Convert.ToInt32(
                        value:
                            await arrayNameCommand.ExecuteScalarAsync(),
                        provider: CultureInfo.InvariantCulture);

                if (existingArrayCount > 0)
                {
                    throw new InvalidOperationException(
                        $"Array '{arrayName}' already exists.");
                }
            }

            #endregion


            #region Define Point Validation Query

            const string pointSql = """
                SELECT
                    PN.[PointName_ID],
                    PN.[Project_ID],
                    PN.[IsDeleted],
                    CASE
                        WHEN CR.[PointName_ID] IS NULL THEN 0
                        ELSE 1
                    END AS [HasReference],
                    CASE
                        WHEN EXISTS
                        (
                            SELECT 1
                            FROM [dbo].[PrismArrayPoint] AS PAP
                                WITH (UPDLOCK, HOLDLOCK)
                            INNER JOIN [dbo].[PrismArray] AS PA
                                WITH (UPDLOCK, HOLDLOCK)
                                ON PA.[Array_ID] = PAP.[Array_ID]
                            WHERE
                                PAP.[PointName_ID] = PN.[PointName_ID]
                                AND PAP.[IsDeleted] = 0
                                AND PA.[IsDeleted] = 0
                                AND PA.[Project_ID] = @Project_ID
                                AND PA.[ArrayType] = @ArrayType
                                AND
                                (
                                    @ExistingArray_ID IS NULL
                                    OR PA.[Array_ID] <> @ExistingArray_ID
                                )
                        ) THEN 1
                        ELSE 0
                    END AS [AlreadyAllocatedToType]
                FROM [dbo].[PointName] AS PN WITH (UPDLOCK, HOLDLOCK)
                LEFT JOIN [dbo].[CoordinatesReference] AS CR WITH (HOLDLOCK)
                    ON CR.[PointName_ID] = PN.[PointName_ID]
                    AND CR.[IsDeleted] = 0
                WHERE PN.[PointName_ID] IN
                (
                    @A_ID,
                    @B_ID,
                    @C_ID,
                    @D_ID,
                    @E_ID
                );
                """;

            #endregion


            #region Execute Point Validation Query

            Dictionary<int, (int ProjectId, bool IsDeleted, bool HasReference, bool AlreadyAllocatedToType)>
                pointStates =
                    new();

            await using SqlCommand pointCommand =
                new(
                    cmdText: pointSql,
                    connection: databaseConnection,
                    transaction: transaction);

            pointCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            pointCommand.Parameters.Add(
                parameterName: "@ArrayType",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    arrayType;

            pointCommand.Parameters.Add(
                parameterName: "@ExistingArray_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    existingArrayId.HasValue
                        ? existingArrayId.Value
                        : DBNull.Value;

            foreach (char pointRole
                in PrismArrayPointRoles)
            {
                SqlParameter pointParameter =
                    pointCommand.Parameters.Add(
                        parameterName: $"@{pointRole}_ID",
                        sqlDbType: System.Data.SqlDbType.Int);

                pointParameter.Value =
                    pointAssignments.TryGetValue(
                        key: pointRole,
                        value: out int pointNameId)
                        ? pointNameId
                        : DBNull.Value;
            }

            await using SqlDataReader reader =
                await pointCommand.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                int pointNameId =
                    reader.GetInt32(
                        i: 0);

                pointStates[pointNameId] =
                    (
                        ProjectId:
                            reader.GetInt32(
                                i: 1),

                        IsDeleted:
                            reader.GetBoolean(
                                i: 2),

                        HasReference:
                            reader.GetInt32(
                                i: 3) == 1,

                        AlreadyAllocatedToType:
                            reader.GetInt32(
                                i: 4) == 1
                    );
            }

            #endregion


            #region Validate Each A-E Point

            foreach (char pointRole
                in GetRequiredPrismArrayPointRoles(
                    arrayType: arrayType))
            {
                int pointNameId =
                    pointAssignments[pointRole];

                if (!pointStates.TryGetValue(
                    key: pointNameId,
                    value: out
                        (int ProjectId, bool IsDeleted, bool HasReference, bool AlreadyAllocatedToType)
                        pointState))
                {
                    throw new InvalidOperationException(
                        $"Point {pointRole} unavailable.");
                }

                if (pointState.ProjectId != projectId)
                {
                    throw new InvalidOperationException(
                        $"Point {pointRole} belongs to another project.");
                }

                if (pointState.IsDeleted)
                {
                    throw new InvalidOperationException(
                        $"Point {pointRole} is deleted.");
                }

                if (!pointState.HasReference)
                {
                    throw new InvalidOperationException(
                        $"Point {pointRole} has no reference coordinates.");
                }

                if (pointState.AlreadyAllocatedToType)
                {
                    throw new InvalidOperationException(
                        $"Point {pointRole} is already allocated to array type {arrayType}.");
                }
            }

            #endregion
        }


        private static async Task<int> InsertPrismArrayAsync(
            int projectId,
            string arrayName,
            int arrayType,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Array Insert

            const string arraySql = """
                INSERT INTO [dbo].[PrismArray]
                (
                    [Project_ID],
                    [ArrayName],
                    [ArrayType],
                    [IsDeleted]
                )
                VALUES
                (
                    @Project_ID,
                    @ArrayName,
                    @ArrayType,
                    0
                );

                SELECT CAST(SCOPE_IDENTITY() AS int);
                """;

            #endregion


            #region Insert Array

            await using SqlCommand arrayCommand =
                new(
                    cmdText: arraySql,
                    connection: databaseConnection,
                    transaction: transaction);

            arrayCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            arrayCommand.Parameters.Add(
                parameterName: "@ArrayName",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 200)
                .Value =
                    arrayName;

            arrayCommand.Parameters.Add(
                parameterName: "@ArrayType",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    arrayType;

            object? arrayIdValue =
                await arrayCommand.ExecuteScalarAsync();

            if (arrayIdValue is null ||
                arrayIdValue == DBNull.Value)
            {
                throw new InvalidOperationException(
                    "Array insert failed.");
            }

            return Convert.ToInt32(
                value: arrayIdValue,
                provider: CultureInfo.InvariantCulture);

            #endregion
        }


        private static async Task InsertPrismArrayPointAsync(
            int arrayId,
            char pointRole,
            int pointNameId,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Array Point Insert

            const string pointSql = """
                INSERT INTO [dbo].[PrismArrayPoint]
                (
                    [Array_ID],
                    [PointRole],
                    [PointName_ID],
                    [IsDeleted]
                )
                VALUES
                (
                    @Array_ID,
                    @PointRole,
                    @PointName_ID,
                    0
                );
                """;

            #endregion


            #region Insert Array Point

            await using SqlCommand pointCommand =
                new(
                    cmdText: pointSql,
                    connection: databaseConnection,
                    transaction: transaction);

            pointCommand.Parameters.Add(
                parameterName: "@Array_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    arrayId;

            pointCommand.Parameters.Add(
                parameterName: "@PointRole",
                sqlDbType: System.Data.SqlDbType.Char,
                size: 1)
                .Value =
                    pointRole.ToString();

            pointCommand.Parameters.Add(
                parameterName: "@PointName_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    pointNameId;

            int rowsInserted =
                await pointCommand.ExecuteNonQueryAsync();

            if (rowsInserted != 1)
            {
                throw new InvalidOperationException(
                    $"Point {pointRole} insert failed.");
            }

            #endregion
        }


        private static async Task UpdatePrismArrayAsync(
            int arrayId,
            int projectId,
            string arrayName,
            int arrayType,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Array Update

            const string arraySql = """
                UPDATE [dbo].[PrismArray]
                SET
                    [ArrayName] = @ArrayName,
                    [ArrayType] = @ArrayType,
                    [IsDeleted] = 0
                WHERE
                    [Array_ID] = @Array_ID
                    AND [Project_ID] = @Project_ID;
                """;

            #endregion


            #region Update Existing Array

            await using SqlCommand arrayCommand =
                new(
                    cmdText: arraySql,
                    connection: databaseConnection,
                    transaction: transaction);

            arrayCommand.Parameters.Add(
                parameterName: "@Array_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value = arrayId;

            arrayCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value = projectId;

            arrayCommand.Parameters.Add(
                parameterName: "@ArrayName",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 200)
                .Value = arrayName;

            arrayCommand.Parameters.Add(
                parameterName: "@ArrayType",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value = arrayType;

            int affectedRows =
                await arrayCommand.ExecuteNonQueryAsync();

            if (affectedRows != 1)
            {
                throw new InvalidOperationException(
                    "The selected array is no longer available for update.");
            }

            #endregion
        }


        private static async Task DeletePrismArrayPointsAsync(
            int arrayId,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Replace Existing Point Assignments

            const string deleteSql = """
                DELETE FROM [dbo].[PrismArrayPoint]
                WHERE [Array_ID] = @Array_ID;
                """;

            await using SqlCommand deleteCommand =
                new(
                    cmdText: deleteSql,
                    connection: databaseConnection,
                    transaction: transaction);

            deleteCommand.Parameters.Add(
                parameterName: "@Array_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value = arrayId;

            await deleteCommand.ExecuteNonQueryAsync();

            #endregion
        }

        #endregion


        #region Committed Array Listing

        private static async Task<List<PrismArraySummaryItem>>
            LoadCommittedPrismArraysAsync(
                int projectId,
                SqlConnection databaseConnection,
                SqlTransaction? transaction)
        {
            #region Define Committed Array Query

            const string arraySql = """
                SELECT
                    PA.[Array_ID],
                    PA.[ArrayType],
                    PAT.[ArrayTypeName],
                    PA.[ArrayName],
                    PA.[IsDeleted]
                FROM [dbo].[PrismArray] AS PA
                INNER JOIN [dbo].[PrismArrayType] AS PAT
                    ON PAT.[ArrayType_ID] = PA.[ArrayType]
                WHERE
                    PA.[Project_ID] = @Project_ID
                ORDER BY
                    PA.[ArrayType],
                    PA.[ArrayName],
                    PA.[Array_ID];
                """;

            #endregion


            #region Execute Committed Array Query

            List<PrismArraySummaryItem> arrays =
                new();

            await using SqlCommand arrayCommand =
                transaction is null
                    ? new SqlCommand(
                        cmdText: arraySql,
                        connection: databaseConnection)
                    : new SqlCommand(
                        cmdText: arraySql,
                        connection: databaseConnection,
                        transaction: transaction);

            arrayCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            await using SqlDataReader reader =
                await arrayCommand.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                arrays.Add(
                    item:
                        new PrismArraySummaryItem
                        {
                            Array_ID =
                                reader.GetInt32(
                                    i: 0),

                            ArrayType =
                                reader.GetInt32(
                                    i: 1),

                            ArrayTypeName =
                                reader.GetString(
                                    i: 2),

                            ArrayName =
                                reader.GetString(
                                    i: 3),

                            IsDeleted =
                                reader.GetBoolean(
                                    i: 4)
                        });
            }

            return arrays;

            #endregion
        }

        #endregion


        #region Array Details

        private async Task<PrismArrayDetailsItem> GetPrismArrayDetailsAsync(
            int arrayId)
        {
            #region Validate Array Context

            if (arrayId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(arrayId),
                    message: "Array_ID must be greater than zero.");
            }

            if (!_activeProjectId.HasValue)
            {
                throw new InvalidOperationException(
                    "Select an active project.");
            }

            int projectId =
                _activeProjectId.Value;

            #endregion


            #region Define Array Detail Query

            const string arraySql = """
                SELECT
                    PA.[Array_ID],
                    PA.[Project_ID],
                    PA.[ArrayType],
                    PAT.[ArrayTypeName],
                    PA.[ArrayName],
                    PA.[IsDeleted],
                    PAP.[PointRole],
                    PN.[PointName_ID],
                    PN.[PointName],
                    PN.[ReplacementName]
                FROM [dbo].[PrismArray] AS PA
                INNER JOIN [dbo].[PrismArrayType] AS PAT
                    ON PAT.[ArrayType_ID] = PA.[ArrayType]
                LEFT JOIN [dbo].[PrismArrayPoint] AS PAP
                    ON PAP.[Array_ID] = PA.[Array_ID]
                    AND PAP.[IsDeleted] = 0
                LEFT JOIN [dbo].[PointName] AS PN
                    ON PN.[PointName_ID] = PAP.[PointName_ID]
                WHERE
                    PA.[Array_ID] = @Array_ID
                    AND PA.[Project_ID] = @Project_ID
                ORDER BY
                    PAP.[PointRole];
                """;

            #endregion


            #region Load Array Details

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand arrayCommand =
                new(
                    cmdText: arraySql,
                    connection: databaseConnection);

            arrayCommand.Parameters.Add(
                parameterName: "@Array_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    arrayId;

            arrayCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            PrismArrayDetailsItem? details =
                null;

            await using SqlDataReader reader =
                await arrayCommand.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                details ??=
                    new PrismArrayDetailsItem
                    {
                        Array_ID =
                            reader.GetInt32(
                                i: 0),

                        Project_ID =
                            reader.GetInt32(
                                i: 1),

                        ArrayType =
                            reader.GetInt32(
                                i: 2),

                        ArrayTypeName =
                            reader.GetString(
                                i: 3),

                        ArrayName =
                            reader.GetString(
                                i: 4),

                        IsDeleted =
                            reader.GetBoolean(
                                i: 5)
                    };

                if (reader.IsDBNull(
                    i: 6) ||
                    reader.IsDBNull(
                        i: 7))
                {
                    continue;
                }

                string roleText =
                    reader.GetString(
                        i: 6);

                if (string.IsNullOrWhiteSpace(roleText))
                {
                    continue;
                }

                char pointRole =
                    ValidatePrismArrayPointRole(
                        pointRole: roleText[0]);

                details.Points.Add(
                    key: pointRole,
                    value:
                        new PrismArrayAvailablePoint
                        {
                            PointName_ID =
                                reader.GetInt32(
                                    i: 7),

                            PointName =
                                reader.GetString(
                                    i: 8),

                            ReplacementName =
                                reader.GetString(
                                    i: 9)
                        });
            }

            #endregion


            #region Validate Array Details

            if (details is null)
            {
                throw new InvalidOperationException(
                    "Array not found.");
            }

            IReadOnlyList<char> requiredPointRoles =
                GetRequiredPrismArrayPointRoles(
                    arrayType: details.ArrayType);

            foreach (char pointRole
                in requiredPointRoles)
            {
                if (!details.Points.ContainsKey(
                    key: pointRole))
                {
                    throw new InvalidOperationException(
                        $"Array point {pointRole} missing.");
                }
            }

            if (details.Points.Count != requiredPointRoles.Count)
            {
                throw new InvalidOperationException(
                    "Array membership invalid.");
            }

            return details;

            #endregion
        }


        private static string BuildPrismArrayDetailsText(
            PrismArrayDetailsItem details)
        {
            #region Build Array Details Text

            ArgumentNullException.ThrowIfNull(
                argument: details);

            StringBuilder message =
                new();

            message.AppendLine(
                value: $"Array ID: {details.Array_ID}");

            message.AppendLine(
                value: $"Array Name: {details.ArrayName}");

            message.AppendLine(
                value: $"Type: {details.ArrayTypeName}");

            message.AppendLine();

            foreach (char pointRole
                in GetRequiredPrismArrayPointRoles(
                    arrayType: details.ArrayType))
            {
                PrismArrayAvailablePoint point =
                    details.Points[pointRole];

                message.AppendLine(
                    value:
                        $"{pointRole}: {point.ReplacementName}");
            }

            return message
                .ToString()
                .TrimEnd();

            #endregion
        }

        #endregion


        #region Array Deletion

        private async Task DeletePrismArrayAsync(
            int arrayId)
        {
            #region Validate Array Delete Context

            if (arrayId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(arrayId),
                    message: "Array_ID must be greater than zero.");
            }

            if (!_activeProjectId.HasValue ||
                string.IsNullOrWhiteSpace(_activeProjectName))
            {
                throw new InvalidOperationException(
                    "Select an active project.");
            }

            int projectId =
                _activeProjectId.Value;

            string projectName =
                _activeProjectName;

            #endregion


            #region Open Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            #endregion


            #region Begin Array Soft-Delete Transaction

            using SqlTransaction transaction =
                databaseConnection.BeginTransaction(
                    iso:
                        System.Data.IsolationLevel.Serializable);

            bool transactionCommitted =
                false;

            #endregion


            try
            {
                #region Acquire Array Definition Lock

                await AcquirePrismArrayDefinitionTransactionLockAsync(
                    projectId: projectId,
                    databaseConnection: databaseConnection,
                    transaction: transaction);

                #endregion


                #region Revalidate Target Project

                await ValidateImportTargetProjectAsync(
                    projectId: projectId,
                    expectedProjectName: projectName,
                    databaseConnection: databaseConnection,
                    transaction: transaction);

                #endregion


                #region Validate Active Array Ownership

                const string arraySql = """
                    SELECT
                        [ArrayName]
                    FROM [dbo].[PrismArray] WITH (UPDLOCK, HOLDLOCK)
                    WHERE
                        [Array_ID] = @Array_ID
                        AND [Project_ID] = @Project_ID
                        AND [IsDeleted] = 0;
                    """;

                await using (SqlCommand arrayCommand =
                    new(
                        cmdText: arraySql,
                        connection: databaseConnection,
                        transaction: transaction))
                {
                    arrayCommand.Parameters.Add(
                        parameterName: "@Array_ID",
                        sqlDbType: System.Data.SqlDbType.Int)
                        .Value =
                            arrayId;

                    arrayCommand.Parameters.Add(
                        parameterName: "@Project_ID",
                        sqlDbType: System.Data.SqlDbType.Int)
                        .Value =
                            projectId;

                    object? arrayNameValue =
                        await arrayCommand.ExecuteScalarAsync();

                    if (arrayNameValue is null ||
                        arrayNameValue == DBNull.Value)
                    {
                        throw new InvalidOperationException(
                            "Array not found.");
                    }
                }

                #endregion


                #region Soft Delete Array Membership

                const string deletePointSql = """
                    UPDATE [dbo].[PrismArrayPoint]
                    SET [IsDeleted] = 1
                    WHERE
                        [Array_ID] = @Array_ID
                        AND [IsDeleted] = 0;
                    """;

                await using (SqlCommand deletePointCommand =
                    new(
                        cmdText: deletePointSql,
                        connection: databaseConnection,
                        transaction: transaction))
                {
                    deletePointCommand.Parameters.Add(
                        parameterName: "@Array_ID",
                        sqlDbType: System.Data.SqlDbType.Int)
                        .Value =
                            arrayId;

                    int deletedPointRows =
                        await deletePointCommand.ExecuteNonQueryAsync();

                    if (deletedPointRows != 5)
                    {
                        throw new InvalidOperationException(
                            "Array membership invalid.");
                    }
                }

                #endregion


                #region Soft Delete Array Definition

                const string deleteArraySql = """
                    UPDATE [dbo].[PrismArray]
                    SET [IsDeleted] = 1
                    WHERE
                        [Array_ID] = @Array_ID
                        AND [Project_ID] = @Project_ID
                        AND [IsDeleted] = 0;
                    """;

                await using (SqlCommand deleteArrayCommand =
                    new(
                        cmdText: deleteArraySql,
                        connection: databaseConnection,
                        transaction: transaction))
                {
                    deleteArrayCommand.Parameters.Add(
                        parameterName: "@Array_ID",
                        sqlDbType: System.Data.SqlDbType.Int)
                        .Value =
                            arrayId;

                    deleteArrayCommand.Parameters.Add(
                        parameterName: "@Project_ID",
                        sqlDbType: System.Data.SqlDbType.Int)
                        .Value =
                            projectId;

                    int deletedArrayRows =
                        await deleteArrayCommand.ExecuteNonQueryAsync();

                    if (deletedArrayRows != 1)
                    {
                        throw new InvalidOperationException(
                            "Array delete failed.");
                    }
                }

                #endregion


                #region Commit Array Soft Delete

                transaction.Commit();

                transactionCommitted =
                    true;

                #endregion


                #region Refresh Array Lists

                if (_prismArrayProjectId.HasValue &&
                    _prismArrayProjectId.Value == projectId)
                {
                    await ReloadPrismArrayConfigurationListsAsync();
                }

                #endregion
            }
            catch
            {
                #region Roll Back Array Soft Delete

                if (!transactionCommitted)
                {
                    try
                    {
                        transaction.Rollback();
                    }
                    catch
                    {
                        // Preserve the original exception.
                    }
                }

                throw;

                #endregion
            }
        }

        #endregion


        #region Tunnel Convergence Chord Helpers

        private static decimal CalculateThreeDimensionalDistance(
            decimal easting1,
            decimal northing1,
            decimal height1,
            decimal easting2,
            decimal northing2,
            decimal height2)
        {
            #region Calculate Three-Dimensional Distance

            double deltaEasting =
                (double)(easting2 - easting1);

            double deltaNorthing =
                (double)(northing2 - northing1);

            double deltaHeight =
                (double)(height2 - height1);

            double distance =
                Math.Sqrt(
                    d:
                        (deltaEasting * deltaEasting) +
                        (deltaNorthing * deltaNorthing) +
                        (deltaHeight * deltaHeight));

            return (decimal)distance;

            #endregion
        }


        private static decimal CalculateTunnelConvergenceChordChange(
            decimal referenceEasting1,
            decimal referenceNorthing1,
            decimal referenceHeight1,
            decimal referenceEasting2,
            decimal referenceNorthing2,
            decimal referenceHeight2,
            decimal currentEasting1,
            decimal currentNorthing1,
            decimal currentHeight1,
            decimal currentEasting2,
            decimal currentNorthing2,
            decimal currentHeight2)
        {
            #region Calculate Reference Chord

            decimal referenceDistance =
                CalculateThreeDimensionalDistance(
                    easting1: referenceEasting1,
                    northing1: referenceNorthing1,
                    height1: referenceHeight1,
                    easting2: referenceEasting2,
                    northing2: referenceNorthing2,
                    height2: referenceHeight2);

            #endregion


            #region Calculate Current Chord

            decimal currentDistance =
                CalculateThreeDimensionalDistance(
                    easting1: currentEasting1,
                    northing1: currentNorthing1,
                    height1: currentHeight1,
                    easting2: currentEasting2,
                    northing2: currentNorthing2,
                    height2: currentHeight2);

            #endregion


            #region Return Signed Chord Change

            return decimal.Round(
                d: currentDistance - referenceDistance,
                decimals: 4,
                mode: MidpointRounding.AwayFromZero);

            #endregion
        }

        #endregion

        #endregion


        #region Prism Array Configuration UI

        #region Arrays Tab Initialisation

        private async void tabConfiguration_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Ignore Child Selection Events

            if (!ReferenceEquals(
                objA: e.OriginalSource,
                objB: sender))
            {
                return;
            }

            #endregion


            #region Update View Array Button State

            bool arraysTabActive =
                tabPrismArrays.IsSelected;

            btnViewPrismArray.Visibility =
                arraysTabActive
                    ? Visibility.Visible
                    : Visibility.Collapsed;

            btnViewPrismArray.IsEnabled =
                arraysTabActive;

            if (!arraysTabActive)
            {
                return;
            }

            #endregion


            #region Load Array Configuration

            try
            {
                await RefreshPrismArrayConfigurationUiAsync();
            }
            catch (Exception ex)
            {
                txtPrismArrayStatus.Text =
                    $"Unable to load arrays: {ex.Message}";
            }

            #endregion
        }


        private async Task RefreshPrismArrayConfigurationUiAsync()
        {
            #region Validate Active Project

            if (!_activeProjectId.HasValue ||
                string.IsNullOrWhiteSpace(_activeProjectName))
            {
                ResetPrismArrayConfigurationState();

                txtPrismArrayStatus.Text =
                    "Select an active project.";

                return;
            }

            #endregion


            #region Establish Or Refresh Array Project Context

            if (!_prismArrayProjectId.HasValue ||
                _prismArrayProjectId.Value != _activeProjectId.Value)
            {
                await InitialisePrismArrayConfigurationForActiveProjectAsync();
            }
            else
            {
                await LoadPrismArrayTypesAsync();

                await ReloadPrismArrayConfigurationListsAsync();
            }

            #endregion


            #region Refresh Array User Interface

            UpdatePrismArrayAssignmentDisplay();

            dgCommittedPrismArrays.SelectedItem =
                null;

            btnViewPrismArray.Visibility =
                Visibility.Visible;

            btnViewPrismArray.IsEnabled =
                true;

            txtPrismArrayStatus.Text =
                $"{_availablePrismArrayPoints.Count} point(s) available; " +
                $"{_committedPrismArrays.Count} array(s) committed.";

            #endregion
        }

        #endregion


        #region Array Definition Controls

        private void cmbPrismArrayType_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Ignore Pre-Initialisation Event

            if (!IsInitialized)
            {
                return;
            }

            #endregion


            #region Remove Assignments Not Used By Selected Type

            if (cmbPrismArrayType.SelectedItem
                is PrismArrayTypeItem selectedArrayType)
            {
                _selectedPrismArrayType =
                    selectedArrayType.ArrayType_ID;

                IReadOnlyList<char> requiredPointRoles =
                    GetRequiredPrismArrayPointRoles(
                        arrayType: selectedArrayType.ArrayType_ID);

                HashSet<char> requiredRoleSet =
                    new(
                        collection: requiredPointRoles);

                List<char> rolesToClear =
                    new();

                foreach (char assignedRole
                    in _prismArrayAssignments.Keys)
                {
                    if (!requiredRoleSet.Contains(
                        item: assignedRole))
                    {
                        rolesToClear.Add(
                            item: assignedRole);
                    }
                }

                foreach (char roleToClear
                    in rolesToClear)
                {
                    ClearPrismArrayPointAssignment(
                        pointRole: roleToClear);
                }
            }

            #endregion


            #region Refresh Role Availability

            UpdatePrismArrayAssignmentDisplay();

            #endregion
        }


        private bool IsPrismArrayPointRoleSelectableFromUi(
            char pointRole)
        {
            #region Resolve Selected Array Type

            if (cmbPrismArrayType.SelectedItem
                is not PrismArrayTypeItem selectedArrayType)
            {
                return true;
            }

            return PrismArrayTypeUsesPointRole(
                arrayType: selectedArrayType.ArrayType_ID,
                pointRole: pointRole);

            #endregion
        }


        private void ClearPrismArrayDefinitionControls()
        {
            #region Clear Definition Header

            txtPrismArrayName.Clear();

            btnCommitPrismArray.Content =
                "Commit Array";

            #endregion


            #region Clear Assignment Display

            txtPrismArrayPointA.Text =
                string.Empty;

            txtPrismArrayPointB.Text =
                string.Empty;

            txtPrismArrayPointC.Text =
                string.Empty;

            txtPrismArrayPointD.Text =
                string.Empty;

            txtPrismArrayPointE.Text =
                string.Empty;

            btnClearPrismArrayA.IsEnabled =
                false;

            btnClearPrismArrayB.IsEnabled =
                false;

            btnClearPrismArrayC.IsEnabled =
                false;

            btnClearPrismArrayD.IsEnabled =
                false;

            btnClearPrismArrayE.IsEnabled =
                false;

            dgPrismArrayAvailablePoints.SelectedItem =
                null;

            #endregion
        }


        private void UpdatePrismArrayAssignmentDisplay()
        {
            #region Resolve Current Assignments

            PrismArrayAvailablePoint? pointA =
                GetPrismArrayPointAssignment(
                    pointRole: 'A');

            PrismArrayAvailablePoint? pointB =
                GetPrismArrayPointAssignment(
                    pointRole: 'B');

            PrismArrayAvailablePoint? pointC =
                GetPrismArrayPointAssignment(
                    pointRole: 'C');

            PrismArrayAvailablePoint? pointD =
                GetPrismArrayPointAssignment(
                    pointRole: 'D');

            PrismArrayAvailablePoint? pointE =
                GetPrismArrayPointAssignment(
                    pointRole: 'E');

            #endregion


            #region Resolve Role Availability

            bool roleASelectable =
                IsPrismArrayPointRoleSelectableFromUi(
                    pointRole: 'A');

            bool roleBSelectable =
                IsPrismArrayPointRoleSelectableFromUi(
                    pointRole: 'B');

            bool roleCSelectable =
                IsPrismArrayPointRoleSelectableFromUi(
                    pointRole: 'C');

            bool roleDSelectable =
                IsPrismArrayPointRoleSelectableFromUi(
                    pointRole: 'D');

            bool roleESelectable =
                IsPrismArrayPointRoleSelectableFromUi(
                    pointRole: 'E');

            #endregion


            #region Update Assignment Text

            txtPrismArrayPointA.Text =
                pointA?.ReplacementName ?? string.Empty;

            txtPrismArrayPointB.Text =
                pointB?.ReplacementName ?? string.Empty;

            txtPrismArrayPointC.Text =
                pointC?.ReplacementName ?? string.Empty;

            txtPrismArrayPointD.Text =
                pointD?.ReplacementName ?? string.Empty;

            txtPrismArrayPointE.Text =
                pointE?.ReplacementName ?? string.Empty;

            #endregion


            #region Update Assignment Controls

            txtPrismArrayPointA.IsEnabled =
                roleASelectable;

            txtPrismArrayPointB.IsEnabled =
                roleBSelectable;

            txtPrismArrayPointC.IsEnabled =
                roleCSelectable;

            txtPrismArrayPointD.IsEnabled =
                roleDSelectable;

            txtPrismArrayPointE.IsEnabled =
                roleESelectable;

            btnAssignPrismArrayA.IsEnabled =
                roleASelectable;

            btnAssignPrismArrayB.IsEnabled =
                roleBSelectable;

            btnAssignPrismArrayC.IsEnabled =
                roleCSelectable;

            btnAssignPrismArrayD.IsEnabled =
                roleDSelectable;

            btnAssignPrismArrayE.IsEnabled =
                roleESelectable;

            btnClearPrismArrayA.IsEnabled =
                roleASelectable &&
                pointA is not null;

            btnClearPrismArrayB.IsEnabled =
                roleBSelectable &&
                pointB is not null;

            btnClearPrismArrayC.IsEnabled =
                roleCSelectable &&
                pointC is not null;

            btnClearPrismArrayD.IsEnabled =
                roleDSelectable &&
                pointD is not null;

            btnClearPrismArrayE.IsEnabled =
                roleESelectable &&
                pointE is not null;

            #endregion
        }


        private void btnClearPrismArrayDefinition_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Clear Pending Array Definition

            ResetPendingPrismArrayDefinition();

            ClearPrismArrayDefinitionControls();

            txtPrismArrayStatus.Text =
                "Definition cleared.";

            #endregion
        }

        #endregion


        #region A-E Assignment Event Handlers

        private void AssignSelectedPrismArrayPointFromUi(
            char pointRole)
        {
            #region Validate Array Type And Role

            int arrayType;

            try
            {
                arrayType =
                    GetSelectedPrismArrayTypeFromUi(
                        arrayTypeComboBox: cmbPrismArrayType);

                if (!PrismArrayTypeUsesPointRole(
                    arrayType: arrayType,
                    pointRole: pointRole))
                {
                    throw new InvalidOperationException(
                        $"Point {char.ToUpperInvariant(pointRole)} is not used by the selected array type.");
                }
            }
            catch (Exception ex)
            {
                txtPrismArrayStatus.Text =
                    ex.Message;

                return;
            }

            #endregion


            #region Validate Selected Available Point

            if (dgPrismArrayAvailablePoints.SelectedItem
                is not PrismArrayAvailablePoint selectedPoint)
            {
                txtPrismArrayStatus.Text =
                    "Select an available point.";

                return;
            }

            #endregion


            #region Assign Point

            try
            {
                AssignPrismArrayPoint(
                    pointRole: pointRole,
                    pointNameId: selectedPoint.PointName_ID);

                dgPrismArrayAvailablePoints.SelectedItem =
                    null;

                UpdatePrismArrayAssignmentDisplay();

                txtPrismArrayStatus.Text =
                    $"Point {char.ToUpperInvariant(pointRole)} assigned.";
            }
            catch (Exception ex)
            {
                txtPrismArrayStatus.Text =
                    ex.Message;
            }

            #endregion
        }


        private void ClearPrismArrayPointFromUi(
            char pointRole)
        {
            #region Clear Point Assignment

            try
            {
                ClearPrismArrayPointAssignment(
                    pointRole: pointRole);

                UpdatePrismArrayAssignmentDisplay();

                txtPrismArrayStatus.Text =
                    $"Point {char.ToUpperInvariant(pointRole)} cleared.";
            }
            catch (Exception ex)
            {
                txtPrismArrayStatus.Text =
                    ex.Message;
            }

            #endregion
        }


        private void btnAssignPrismArrayA_Click(object sender, RoutedEventArgs e) =>
            AssignSelectedPrismArrayPointFromUi(pointRole: 'A');

        private void btnAssignPrismArrayB_Click(object sender, RoutedEventArgs e) =>
            AssignSelectedPrismArrayPointFromUi(pointRole: 'B');

        private void btnAssignPrismArrayC_Click(object sender, RoutedEventArgs e) =>
            AssignSelectedPrismArrayPointFromUi(pointRole: 'C');

        private void btnAssignPrismArrayD_Click(object sender, RoutedEventArgs e) =>
            AssignSelectedPrismArrayPointFromUi(pointRole: 'D');

        private void btnAssignPrismArrayE_Click(object sender, RoutedEventArgs e) =>
            AssignSelectedPrismArrayPointFromUi(pointRole: 'E');


        private void btnClearPrismArrayA_Click(object sender, RoutedEventArgs e) =>
            ClearPrismArrayPointFromUi(pointRole: 'A');

        private void btnClearPrismArrayB_Click(object sender, RoutedEventArgs e) =>
            ClearPrismArrayPointFromUi(pointRole: 'B');

        private void btnClearPrismArrayC_Click(object sender, RoutedEventArgs e) =>
            ClearPrismArrayPointFromUi(pointRole: 'C');

        private void btnClearPrismArrayD_Click(object sender, RoutedEventArgs e) =>
            ClearPrismArrayPointFromUi(pointRole: 'D');

        private void btnClearPrismArrayE_Click(object sender, RoutedEventArgs e) =>
            ClearPrismArrayPointFromUi(pointRole: 'E');

        #endregion


        #region Array Commit UI

        private static int GetSelectedPrismArrayTypeFromUi(
            ComboBox arrayTypeComboBox)
        {
            #region Resolve Selected Database Array Type

            ArgumentNullException.ThrowIfNull(
                argument: arrayTypeComboBox);

            if (arrayTypeComboBox.SelectedItem
                is not PrismArrayTypeItem selectedArrayType)
            {
                throw new InvalidOperationException(
                    "Select an array type.");
            }

            return selectedArrayType.ArrayType_ID;

            #endregion
        }


        private async void btnCommitPrismArray_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Prepare Array Definition

            string arrayName =
                txtPrismArrayName.Text?.Trim()
                ?? string.Empty;

            int arrayType;

            try
            {
                if (!_activeProjectId.HasValue)
                {
                    throw new InvalidOperationException(
                        "Select an active project.");
                }

                if (!_prismArrayProjectId.HasValue ||
                    _prismArrayProjectId.Value != _activeProjectId.Value)
                {
                    await InitialisePrismArrayConfigurationForActiveProjectAsync();
                }

                arrayType =
                    GetSelectedPrismArrayTypeFromUi(
                        arrayTypeComboBox: cmbPrismArrayType);

                SetPendingPrismArrayDefinition(
                    arrayType: arrayType,
                    arrayName: arrayName);

                _ = CapturePrismArrayPointAssignments(
                    arrayType: arrayType);
            }
            catch (Exception ex)
            {
                txtPrismArrayStatus.Text =
                    ex.Message;

                return;
            }

            #endregion


            #region Confirm Array Commit

            StringBuilder confirmationBuilder =
                new();

            confirmationBuilder.AppendLine(
                value: $"Project: {_prismArrayProjectName}");

            confirmationBuilder.AppendLine(
                value: $"Type: {GetPrismArrayType(arrayTypeId: arrayType).ArrayTypeName}");

            confirmationBuilder.AppendLine(
                value: $"Array: {arrayName}");

            confirmationBuilder.AppendLine();

            foreach (char pointRole
                in GetRequiredPrismArrayPointRoles(
                    arrayType: arrayType))
            {
                confirmationBuilder.AppendLine(
                    value:
                        $"{pointRole}: " +
                        $"{GetPrismArrayPointAssignment(pointRole: pointRole)!.ReplacementName}");
            }

            confirmationBuilder.AppendLine();

            confirmationBuilder.Append(
                value:
                    _editingPrismArrayId.HasValue
                        ? "Update this array?"
                        : "Commit this array?");

            string confirmationMessage =
                confirmationBuilder.ToString();

            MessageBoxResult confirmation =
                MessageBox.Show(
                    owner: this,
                    messageBoxText: confirmationMessage,
                    caption:
                        _editingPrismArrayId.HasValue
                            ? "Confirm Array Update"
                            : "Confirm Array",
                    button: MessageBoxButton.YesNo,
                    icon: MessageBoxImage.Question,
                    defaultResult: MessageBoxResult.No);

            if (confirmation != MessageBoxResult.Yes)
            {
                txtPrismArrayStatus.Text =
                    "Commit cancelled.";

                return;
            }

            #endregion


            #region Commit Array

            btnCommitPrismArray.IsEnabled =
                false;

            try
            {
                bool updatingExistingArray =
                    _editingPrismArrayId.HasValue;

                int arrayId =
                    await CommitCurrentPrismArrayDefinitionAsync();

                txtPrismArrayName.Clear();

                btnCommitPrismArray.Content =
                    "Commit Array";

                UpdatePrismArrayAssignmentDisplay();

                foreach (PrismArraySummaryItem committedArray
                    in _committedPrismArrays)
                {
                    if (committedArray.Array_ID == arrayId)
                    {
                        dgCommittedPrismArrays.SelectedItem =
                            committedArray;

                        dgCommittedPrismArrays.ScrollIntoView(
                            item: committedArray);

                        break;
                    }
                }

                txtPrismArrayStatus.Text =
                    updatingExistingArray
                        ? $"Array '{arrayName}' updated."
                        : $"Array '{arrayName}' committed.";
            }
            catch (Exception ex)
            {
                txtPrismArrayStatus.Text =
                    $"Commit failed: {ex.Message}";
            }
            finally
            {
                btnCommitPrismArray.IsEnabled =
                    true;
            }

            #endregion
        }

        #endregion


        #region Existing Array Update

        private void dgCommittedPrismArrays_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Maintain View Button State

            btnViewPrismArray.IsEnabled =
                tabPrismArrays.IsSelected;

            #endregion
        }


        private async void PrismArrayIncludedCheckBox_Click(
            object sender,
            RoutedEventArgs e)
        {
            if (sender is not CheckBox checkBox ||
                checkBox.DataContext is not PrismArraySummaryItem array)
            {
                return;
            }

            bool includeArray =
                checkBox.IsChecked == true;

            const string sql = """
                UPDATE [dbo].[PrismArray]
                SET [IsDeleted] = @IsDeleted
                WHERE
                    [Array_ID] = @Array_ID
                    AND [Project_ID] = @Project_ID;
                """;

            try
            {
                await using SqlConnection databaseConnection =
                    new(
                        connectionString: GetTrackGeometryConnectionString());

                await databaseConnection.OpenAsync();

                await using SqlCommand command =
                    new(
                        cmdText: sql,
                        connection: databaseConnection);

                command.Parameters.Add(
                    parameterName: "@IsDeleted",
                    sqlDbType: System.Data.SqlDbType.Bit)
                    .Value = !includeArray;

                command.Parameters.Add(
                    parameterName: "@Array_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value = array.Array_ID;

                command.Parameters.Add(
                    parameterName: "@Project_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value = _activeProjectId
                        ?? throw new InvalidOperationException(
                            "Select an active project.");

                int affectedRows =
                    await command.ExecuteNonQueryAsync();

                if (affectedRows != 1)
                {
                    throw new InvalidOperationException(
                        "The array changed before its inclusion state could be saved.");
                }

                array.IsDeleted =
                    !includeArray;

                txtPrismArrayStatus.Text =
                    includeArray
                        ? $"Array '{array.ArrayName}' included."
                        : $"Array '{array.ArrayName}' excluded.";
            }
            catch (Exception ex)
            {
                checkBox.IsChecked =
                    !includeArray;

                array.IsDeleted =
                    includeArray;

                txtPrismArrayStatus.Text =
                    $"Unable to update array inclusion: {ex.Message}";
            }
        }


        private async void btnViewPrismArray_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Array Review Context

            if (!tabPrismArrays.IsSelected)
            {
                return;
            }

            if (!_activeProjectId.HasValue ||
                string.IsNullOrWhiteSpace(_activeProjectName))
            {
                txtPrismArrayStatus.Text =
                    "Select an active project.";

                return;
            }

            #endregion


            #region Prepare Array Review

            btnViewPrismArray.IsEnabled =
                false;

            #endregion


            try
            {
                #region Reload Committed Arrays From Database

                await RefreshPrismArrayConfigurationUiAsync();

                if (_committedPrismArrays.Count == 0)
                {
                    MessageBox.Show(
                        owner: this,
                        messageBoxText:
                            "No committed arrays.",
                        caption:
                            "Update Array",
                        button:
                            MessageBoxButton.OK,
                        icon:
                            MessageBoxImage.Information);

                    txtPrismArrayStatus.Text =
                        "No committed arrays.";

                    return;
                }

                #endregion


                #region Build Array Selection List

                List<ArraySelectionItem> arraySelectionItems =
                    new();

                foreach (PrismArraySummaryItem committedArray
                    in _committedPrismArrays)
                {
                    arraySelectionItems.Add(
                        item:
                            new ArraySelectionItem
                            {
                                Array_ID =
                                    committedArray.Array_ID,

                                ArrayTypeName =
                                    committedArray.ArrayTypeName,

                                ArrayName =
                                    committedArray.ArrayName
                            });
                }

                #endregion


                #region Select Committed Array

                ArraySelectionWindow selectionWindow =
                    new(
                        projectName:
                            _activeProjectName,
                        arrays:
                            arraySelectionItems)
                    {
                        Owner =
                            this
                    };

                bool? selectionResult =
                    selectionWindow.ShowDialog();

                if (selectionResult != true ||
                    !selectionWindow.SelectedArrayId.HasValue)
                {
                    txtPrismArrayStatus.Text =
                        "Array update cancelled.";

                    return;
                }

                int selectedArrayId =
                    selectionWindow.SelectedArrayId.Value;

                #endregion


                #region Load Selected Array Details

                PrismArrayDetailsItem details =
                    await GetPrismArrayDetailsAsync(
                        arrayId:
                            selectedArrayId);

                #endregion


                #region Load Array Into Definition Editor

                ResetPendingPrismArrayDefinition();

                await ReloadPrismArrayConfigurationListsAsync();

                cmbPrismArrayType.SelectedValue =
                    details.ArrayType;

                txtPrismArrayName.Text =
                    details.ArrayName;

                SetPendingPrismArrayDefinition(
                    arrayType: details.ArrayType,
                    arrayName: details.ArrayName);

                _editingPrismArrayId =
                    details.Array_ID;

                foreach (char pointRole
                    in GetRequiredPrismArrayPointRoles(
                        arrayType: details.ArrayType))
                {
                    AssignPrismArrayPoint(
                        pointRole: pointRole,
                        pointNameId:
                            details.Points[pointRole].PointName_ID);
                }

                btnCommitPrismArray.Content =
                    "Update Array";

                UpdatePrismArrayAssignmentDisplay();

                txtPrismArrayStatus.Text =
                    $"Array '{details.ArrayName}' loaded for update.";

                #endregion
            }
            catch (Exception ex)
            {
                txtPrismArrayStatus.Text =
                    ex.Message;
            }
            finally
            {
                #region Restore View Array Button

                btnViewPrismArray.Visibility =
                    tabPrismArrays.IsSelected
                        ? Visibility.Visible
                        : Visibility.Collapsed;

                btnViewPrismArray.IsEnabled =
                    tabPrismArrays.IsSelected;

                #endregion
            }
        }

        #endregion

        #endregion


        #region Constructor

        private static IReadOnlyList<TimeZoneInfo> GetInitialProjectTimeZones()
        {
            HashSet<string> supportedTimeZoneIds =
                new(
                    collection:
                    [
                        "GMT Standard Time",
                        "W. Europe Standard Time",
                        "Central Europe Standard Time",
                        "Central European Standard Time",
                        "Romance Standard Time",
                        "E. Europe Standard Time",
                        "FLE Standard Time",
                        "GTB Standard Time",
                        "AUS Eastern Standard Time",
                        "E. Australia Standard Time",
                        "Cen. Australia Standard Time",
                        "W. Australia Standard Time",
                        "Tasmania Standard Time",
                        "Lord Howe Standard Time"
                    ],
                    comparer: StringComparer.OrdinalIgnoreCase);

            return TimeZoneInfo.GetSystemTimeZones()
                .Where(
                    predicate:
                        timeZone => supportedTimeZoneIds.Contains(
                            item: timeZone.Id))
                .OrderBy(
                    keySelector: timeZone => timeZone.BaseUtcOffset)
                .ThenBy(
                    keySelector: timeZone => timeZone.DisplayName,
                    comparer: StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        public MainWindow()
        {
            #region Initialise Window Components

            InitializeComponent();

            _chartPreviewRenderer =
                new WpfCanvasChartPreviewRenderer(
                    owner: this);

            cmbProjectTimeZone.ItemsSource =
                TimeZoneInfo.GetSystemTimeZones()
                    .OrderBy(
                        keySelector: timeZone => timeZone.BaseUtcOffset)
                    .ThenBy(
                        keySelector: timeZone => timeZone.DisplayName,
                        comparer: StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

            cmbProjectTimeZone.SelectedValue =
                TimeZoneInfo.Utc.Id;

            UpdateConfigurationWorkflowTabAvailability();

            #endregion


            #region EPPlus License

            gnaT.epplusLicense();

            #endregion


            #region Initialise Project DataGrid

            dgProjects.ItemsSource =
                _projectItems;

            _pendingProjectStartDate =
                DateTime.Today;

            dpProjectStartDate.SelectedDate =
                _pendingProjectStartDate;

            #endregion


            #region Initialise Chart Configuration

            InitialiseChartConfigurationUi();

            cmbExistingChart.ItemsSource =
                _existingCharts;

            cmbChartTemplate.ItemsSource =
                _chartTemplates;

            #endregion


            #region Initialise Reference Coordinate Import DataGrid

            dgReferenceCoordinateImport.ItemsSource =
                _referenceImportItems;

            #endregion


            #region Initialise Geotechnical Sensor Import DataGrid

            dgGeotechSensorImport.ItemsSource =
                _geotechImportItems;

            #endregion


            #region Initialise Prism Array DataGrids

            dgPrismArrayAvailablePoints.ItemsSource =
                _availablePrismArrayPoints;

            dgCommittedPrismArrays.ItemsSource =
                _committedPrismArrays;

            #endregion


            #region Register Window Events

            Loaded +=
                MainWindow_Loaded;

            Closed +=
                MainWindow_Closed;

            #endregion
        }

        #endregion


        #region Configuration Workflow Availability

        private void UpdateConfigurationWorkflowTabAvailability()
        {
            #region Resolve Active Project State

            bool activeProjectAvailable =
                _activeProjectId.HasValue &&
                !string.IsNullOrWhiteSpace(
                    value: _activeProjectName);

            #endregion


            #region Apply Workflow Tab State

            tabConfiguration.IsEnabled =
                activeProjectAvailable;

            tabReferenceCoordinates.IsEnabled =
                activeProjectAvailable;

            tabPrismPairs.IsEnabled =
                activeProjectAvailable;

            tabPrismArrays.IsEnabled =
                activeProjectAvailable;

            tabGeotech.IsEnabled =
                activeProjectAvailable;

            tabCharts.IsEnabled =
                activeProjectAvailable;

            #endregion
        }

        #endregion


        #region Window Initialisation

        #region Application Startup

        private async void MainWindow_Loaded(
            object sender,
            RoutedEventArgs e)
        {
            #region Load Persisted Startup Configuration

            // Registry values are startup defaults only.
            // They are read once when this process starts.

            LoadDatabaseConnectionString();

            LoadActiveProjectFromRegistry();

            #endregion


            #region Initialise Process-Local Active Project

            // A persisted Project_ID does not become operational merely because
            // it exists in the Registry.
            //
            // The project must first be validated against SQL Server and this
            // application instance must acquire its Shared project lock.

            await InitialiseStartupActiveProjectAsync();

            UpdateConfigurationWorkflowTabAvailability();

            if (_activeProjectId.HasValue)
            {
                try
                {
                    await EnsureChartTypeCatalogueAsync();

                    await EnsureChartTemplateTablesAsync();

                    await RefreshExistingChartsAsync();

                    await RefreshChartTemplatesAsync();
                }
                catch (Exception chartEx)
                {
                    txtDbConnectionStatus.Text =
                        $"Chart catalogue unavailable: {chartEx.Message}";
                }
            }

            #endregion
        }

        #endregion



        private void MainWindow_Closed(
            object? sender,
            EventArgs e)
        {
            #region Release Cross-Instance Project Lock

            ReleaseActiveProjectLockConnection();

            #endregion
        }




        #region Registry Persistence

        #region Database Connection String Persistence

        private void LoadDatabaseConnectionString()
        {
            #region Initialise Connection String

            string connectionString =
                string.Empty;

            #endregion


            #region Read Connection String From Registry

            using RegistryKey? registryKey =
                Registry.CurrentUser.OpenSubKey(
                    name: RegistryPath,
                    writable: false);

            if (registryKey is not null)
            {
                connectionString =
                    registryKey.GetValue(
                        name: RegistryDatabaseConnectionString,
                        defaultValue: string.Empty)?.ToString()
                    ?? string.Empty;
            }

            #endregion


            #region Populate Connection String Control

            txtDbConnectionString.Text =
                connectionString;

            _validatedDatabaseConnectionString =
                connectionString;

            #endregion
        }


        private static void SaveDatabaseConnectionString(
            string connectionString)
        {
            #region Validate Connection String

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new ArgumentException(
                    message: "Database connection string cannot be empty.",
                    paramName: nameof(connectionString));
            }

            #endregion


            #region Write Connection String To Registry

            using RegistryKey registryKey =
                Registry.CurrentUser.CreateSubKey(
                    subkey: RegistryPath,
                    writable: true)
                ?? throw new InvalidOperationException(
                    $"Unable to create or open registry key " +
                    $"'HKEY_CURRENT_USER\\{RegistryPath}'.");

            registryKey.SetValue(
                name: RegistryDatabaseConnectionString,
                value: connectionString,
                valueKind: RegistryValueKind.String);

            #endregion
        }

        #endregion


        #region Active Project Persistence

        private void LoadActiveProjectFromRegistry()
        {
            #region Initialise Process-Local Active Project

            _activeProjectId = null;

            _activeProjectName =
                string.Empty;

            #endregion


            #region Read Persisted Startup Default

            using RegistryKey? registryKey =
                Registry.CurrentUser.OpenSubKey(
                    name: RegistryPath,
                    writable: false);

            object? activeProjectIdValue =
                registryKey?.GetValue(
                    name: RegistryActiveProjectId,
                    defaultValue: null);

            string activeProjectName =
                registryKey?.GetValue(
                    name: RegistryActiveProjectName,
                    defaultValue: string.Empty)?.ToString()
                ?? string.Empty;

            #endregion


            #region Populate Process-Local Runtime State

            if (activeProjectIdValue is int activeProjectId &&
                activeProjectId > 0)
            {
                _activeProjectId =
                    activeProjectId;

                _activeProjectName =
                    activeProjectName.Trim();
            }

            #endregion


            #region Update Active Project Display

            txtActiveProject.Text =
                _activeProjectId.HasValue &&
                !string.IsNullOrWhiteSpace(_activeProjectName)
                    ? _activeProjectName
                    : "No active project";

            #endregion
        }


        private static void SaveActiveProjectToRegistry(
            int projectId,
            string projectName)
        {
            #region Validate Active Project

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(projectId),
                    message: "Active Project_ID must be greater than zero.");
            }

            string validatedProjectName =
                projectName?.Trim()
                ?? throw new ArgumentNullException(
                    paramName: nameof(projectName));

            if (string.IsNullOrWhiteSpace(validatedProjectName))
            {
                throw new ArgumentException(
                    message: "Active project name cannot be empty.",
                    paramName: nameof(projectName));
            }

            #endregion


            #region Persist Startup Default

            using RegistryKey registryKey =
                Registry.CurrentUser.CreateSubKey(
                    subkey: RegistryPath,
                    writable: true)
                ?? throw new InvalidOperationException(
                    $"Unable to create or open registry key " +
                    $"'HKEY_CURRENT_USER\\{RegistryPath}'.");

            registryKey.SetValue(
                name: RegistryActiveProjectId,
                value: projectId,
                valueKind: RegistryValueKind.DWord);

            registryKey.SetValue(
                name: RegistryActiveProjectName,
                value: validatedProjectName,
                valueKind: RegistryValueKind.String);

            #endregion
        }


        private static void ClearActiveProjectFromRegistry()
        {
            #region Open Registry Configuration

            using RegistryKey? registryKey =
                Registry.CurrentUser.OpenSubKey(
                    name: RegistryPath,
                    writable: true);

            if (registryKey is null)
            {
                return;
            }

            #endregion


            #region Remove Persisted Startup Default

            registryKey.DeleteValue(
                name: RegistryActiveProjectId,
                throwOnMissingValue: false);

            registryKey.DeleteValue(
                name: RegistryActiveProjectName,
                throwOnMissingValue: false);

            #endregion
        }

        #endregion

        #endregion

        #endregion


        #region Database Connection Infrastructure

        private string GetValidatedDatabaseConnectionString()
        {
            #region Validate Operational Connection String

            if (string.IsNullOrWhiteSpace(_validatedDatabaseConnectionString))
            {
                throw new InvalidOperationException(
                    "Test the database connection first.");
            }

            string enteredConnectionString =
                txtDbConnectionString.Text?.Trim()
                ?? string.Empty;

            if (!ConnectionStringsEquivalent(
                firstConnectionString: enteredConnectionString,
                secondConnectionString: _validatedDatabaseConnectionString))
            {
                throw new InvalidOperationException(
                    "Database connection changed. Test it first.");
            }

            return _validatedDatabaseConnectionString;

            #endregion
        }


        private string GetTrackGeometryConnectionString()
        {
            #region Build Track Geometry Connection String

            return BuildDatabaseConnectionString(
                baseConnectionString: GetValidatedDatabaseConnectionString(),
                databaseName: TrackGeometryDatabaseName,
                pooling: null);

            #endregion
        }


        private static string BuildDatabaseConnectionString(
            string baseConnectionString,
            string databaseName,
            bool? pooling)
        {
            #region Build Database Connection String

            if (string.IsNullOrWhiteSpace(baseConnectionString))
            {
                throw new ArgumentException(
                    message: "Database connection string cannot be empty.",
                    paramName: nameof(baseConnectionString));
            }

            SqlConnectionStringBuilder connectionBuilder =
                new(
                    connectionString: baseConnectionString)
                {
                    InitialCatalog =
                        databaseName
                };

            if (pooling.HasValue)
            {
                connectionBuilder.Pooling =
                    pooling.Value;
            }

            return connectionBuilder.ConnectionString;

            #endregion
        }


        private static bool ConnectionStringsEquivalent(
            string firstConnectionString,
            string secondConnectionString)
        {
            #region Compare Normalised Connection Strings

            if (string.IsNullOrWhiteSpace(firstConnectionString) ||
                string.IsNullOrWhiteSpace(secondConnectionString))
            {
                return string.Equals(
                    a: firstConnectionString?.Trim(),
                    b: secondConnectionString?.Trim(),
                    comparisonType: StringComparison.Ordinal);
            }

            string firstNormalised =
                new SqlConnectionStringBuilder(
                    connectionString: firstConnectionString)
                    .ConnectionString;

            string secondNormalised =
                new SqlConnectionStringBuilder(
                    connectionString: secondConnectionString)
                    .ConnectionString;

            return string.Equals(
                a: firstNormalised,
                b: secondNormalised,
                comparisonType: StringComparison.Ordinal);

            #endregion
        }


        private void ResetDatabaseDependentRuntimeState(
            bool clearPersistedActiveProject)
        {
            #region Release Active Project Lock

            ReleaseActiveProjectLockConnection();

            #endregion


            #region Clear Process State

            _activeProjectId =
                null;

            _activeProjectName =
                string.Empty;

            txtActiveProject.Text =
                "No active project";

            _projectItems.Clear();

            ResetReferenceImportState(
                statusMessage: "No CSV selected.");

            ResetGeotechImportState(
                statusMessage: "No CSV selected.");

            ResetPrismPairImportState(
                statusMessage: "No workbook selected.");

            ResetPrismArrayConfigurationState();

            UpdateConfigurationWorkflowTabAvailability();

            #endregion


            #region Clear Persisted Active Project

            if (clearPersistedActiveProject)
            {
                ClearActiveProjectFromRegistry();
            }

            #endregion
        }


        private async void btnTestDbConnection_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? string.Empty;

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                txtDbConnectionStatus.Text =
                    "Enter a database connection string.";

                return;
            }

            #endregion


            #region Prepare Connection Test

            btnTestDbConnection.IsEnabled =
                false;

            txtDbConnectionStatus.Text =
                "Testing database connection...";

            #endregion


            try
            {
                #region Test SQL Database Connection

                await using SqlConnection sqlConnection =
                    new(
                        connectionString: connectionString);

                await sqlConnection.OpenAsync();

                #endregion


                #region Handle Connection Change

                bool connectionChanged =
                    !string.IsNullOrWhiteSpace(_validatedDatabaseConnectionString) &&
                    !ConnectionStringsEquivalent(
                        firstConnectionString: connectionString,
                        secondConnectionString: _validatedDatabaseConnectionString);

                if (connectionChanged)
                {
                    ResetDatabaseDependentRuntimeState(
                        clearPersistedActiveProject: true);
                }

                _validatedDatabaseConnectionString =
                    connectionString;

                SaveDatabaseConnectionString(
                    connectionString: connectionString);

                #endregion


                #region Report Success

                txtDbConnectionStatus.Text =
                    $"Connection successful. " +
                    $"Server: {sqlConnection.DataSource}; " +
                    $"Database: {sqlConnection.Database}";

                #endregion
            }
            catch (SqlException ex)
            {
                txtDbConnectionStatus.Text =
                    $"SQL connection failed: {ex.Message}";
            }
            catch (ArgumentException ex)
            {
                txtDbConnectionStatus.Text =
                    $"Invalid connection: {ex.Message}";
            }
            catch (InvalidOperationException ex)
            {
                txtDbConnectionStatus.Text =
                    ex.Message;
            }
            catch (Exception ex)
            {
                txtDbConnectionStatus.Text =
                    $"Connection test failed: {ex.Message}";
            }
            finally
            {
                btnTestDbConnection.IsEnabled =
                    true;
            }
        }

        #endregion


        #region Import Target Project Validation

        private static async Task ValidateImportTargetProjectAsync(
            int projectId,
            string expectedProjectName,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Read Target Project

            const string projectSql = """
                SELECT
                    [ProjectName]
                FROM [dbo].[Project] WITH (UPDLOCK, HOLDLOCK)
                WHERE
                    [Project_ID] = @Project_ID
                    AND [IsDeleted] = 0;
                """;

            await using SqlCommand projectCommand =
                new(
                    cmdText: projectSql,
                    connection: databaseConnection,
                    transaction: transaction);

            projectCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            object? projectNameValue =
                await projectCommand.ExecuteScalarAsync();

            if (projectNameValue is null ||
                projectNameValue == DBNull.Value)
            {
                throw new InvalidOperationException(
                    "Target project unavailable.");
            }

            string databaseProjectName =
                Convert.ToString(
                    value: projectNameValue,
                    provider: System.Globalization.CultureInfo.InvariantCulture)
                ?? throw new InvalidOperationException(
                    "Target project name unavailable.");

            #endregion


            #region Verify Project Name

            if (!string.Equals(
                a: databaseProjectName,
                b: expectedProjectName,
                comparisonType: StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Target project changed. Reselect the import source.");
            }

            #endregion
        }

        #endregion


        #region Database Creation

        private async void btnCreateDb_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Connection String

            string connectionString;

            try
            {
                connectionString =
                    GetValidatedDatabaseConnectionString();
            }
            catch (InvalidOperationException ex)
            {
                txtDbConnectionStatus.Text =
                    ex.Message;

                return;
            }

            #endregion


            #region Prepare User Interface

            btnCreateDb.IsEnabled =
                false;

            btnTestDbConnection.IsEnabled =
                false;

            txtDbConnectionStatus.Text =
                $"Checking database '{TrackGeometryDatabaseName}'...";

            #endregion


            try
            {
                #region Check Whether Database Exists

                bool databaseExists =
                    await DatabaseExistsAsync(
                        connectionString: connectionString);

                #endregion


                #region Confirm Existing Database Recreation

                bool recreateExistingDatabase =
                    false;

                if (databaseExists)
                {
                    #region First Recreation Warning

                    MessageBoxResult confirmation =
                        MessageBox.Show(
                            messageBoxText:
                                $"Database '{TrackGeometryDatabaseName}' already exists.\n\n" +
                                "Recreating the database will permanently delete ALL " +
                                "existing data and completely recreate the database and " +
                                "all table structures.\n\n" +
                                "THIS OPERATION CANNOT BE UNDONE.\n\n" +
                                "Do you want to permanently delete and recreate the database?",
                            caption: "Recreate Database Warning",
                            button: MessageBoxButton.YesNo,
                            icon: MessageBoxImage.Warning,
                            defaultResult: MessageBoxResult.No);

                    if (confirmation != MessageBoxResult.Yes)
                    {
                        txtDbConnectionStatus.Text =
                            $"Database '{TrackGeometryDatabaseName}' was not changed.";

                        return;
                    }

                    #endregion


                    #region Final Recreation Warning

                    ConfirmDatabaseRecreationWindow finalConfirmation =
                        new()
                        {
                            Owner = this
                        };

                    bool? proceedWithRecreation =
                        finalConfirmation.ShowDialog();

                    if (proceedWithRecreation != true)
                    {
                        txtDbConnectionStatus.Text =
                            $"Database '{TrackGeometryDatabaseName}' recreation was aborted.";

                        return;
                    }

                    #endregion


                    #region Set Recreation Mode

                    recreateExistingDatabase =
                        true;

                    #endregion
                }

                #endregion


                #region Create Or Recreate Database

                txtDbConnectionStatus.Text =
                    recreateExistingDatabase
                        ? $"Recreating database '{TrackGeometryDatabaseName}'..."
                        : $"Creating database '{TrackGeometryDatabaseName}'...";

                if (recreateExistingDatabase)
                {
                    ResetDatabaseDependentRuntimeState(
                        clearPersistedActiveProject: true);
                }

                await CreateDatabaseAndTablesAsync(
                    connectionString: connectionString,
                    recreateExistingDatabase: recreateExistingDatabase);

                #endregion


                #region Report Successful Completion

                txtDbConnectionStatus.Text =
                    recreateExistingDatabase
                        ? $"Database '{TrackGeometryDatabaseName}' was recreated successfully. " +
                          $"All tables are empty."
                        : $"Database '{TrackGeometryDatabaseName}' and all required tables " +
                          $"were created successfully.";

                #endregion
            }

            #region Handle Database Errors

            catch (SqlException ex)
            {
                txtDbConnectionStatus.Text =
                    $"Database operation failed: {ex.Message}";
            }
            catch (ArgumentException ex)
            {
                txtDbConnectionStatus.Text =
                    $"Invalid database configuration: {ex.Message}";
            }
            catch (InvalidOperationException ex)
            {
                txtDbConnectionStatus.Text =
                    $"Database operation failed: {ex.Message}";
            }
            catch (Exception ex)
            {
                txtDbConnectionStatus.Text =
                    $"Database operation failed: {ex.Message}";
            }

            #endregion


            #region Restore User Interface

            finally
            {
                btnCreateDb.IsEnabled =
                    true;

                btnTestDbConnection.IsEnabled =
                    true;
            }

            #endregion
        }


        private static async Task<bool> DatabaseExistsAsync(
            string connectionString)
        {
            #region Validate Connection String

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new ArgumentException(
                    message: "Database connection string cannot be empty.",
                    paramName: nameof(connectionString));
            }

            #endregion


            #region Build Master Database Connection String

            string masterConnectionString =
                BuildDatabaseConnectionString(
                    baseConnectionString: connectionString,
                    databaseName: "master",
                    pooling: null);

            #endregion


            #region Check Database Existence

            const string databaseExistsSql = """
                SELECT
                    CASE
                        WHEN DB_ID(N'DBTrackGeometry') IS NULL THEN 0
                        ELSE 1
                    END;
                """;

            await using SqlConnection masterConnection =
                new(
                    connectionString: masterConnectionString);

            await masterConnection.OpenAsync();

            await using SqlCommand databaseExistsCommand =
                new(
                    cmdText: databaseExistsSql,
                    connection: masterConnection);

            object databaseExistsResult =
                await databaseExistsCommand.ExecuteScalarAsync()
                ?? throw new InvalidOperationException(
                    "SQL Server returned no result when checking database existence.");

            return Convert.ToInt32(
                value: databaseExistsResult) == 1;

            #endregion
        }


        private static async Task CreateDatabaseAndTablesAsync(
            string connectionString,
            bool recreateExistingDatabase)
        {
            #region Validate Connection String

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new ArgumentException(
                    message: "Database connection string cannot be empty.",
                    paramName: nameof(connectionString));
            }

            #endregion


            #region Build Master Database Connection String

            string masterConnectionString =
                BuildDatabaseConnectionString(
                    baseConnectionString: connectionString,
                    databaseName: "master",
                    pooling: null);

            #endregion


            #region Create Or Recreate Database

            await using (SqlConnection masterConnection =
                new(
                    connectionString: masterConnectionString))
            {
                await masterConnection.OpenAsync();

                string databaseCreationSql;

                if (recreateExistingDatabase)
                {
                    databaseCreationSql = """
                        IF DB_ID(N'DBTrackGeometry') IS NOT NULL
                        BEGIN
                            ALTER DATABASE [DBTrackGeometry]
                            SET SINGLE_USER
                            WITH ROLLBACK IMMEDIATE;

                            DROP DATABASE [DBTrackGeometry];
                        END;

                        CREATE DATABASE [DBTrackGeometry];
                        """;
                }
                else
                {
                    databaseCreationSql = """
                        IF DB_ID(N'DBTrackGeometry') IS NULL
                        BEGIN
                            CREATE DATABASE [DBTrackGeometry];
                        END;
                        """;
                }

                await using SqlCommand createDatabaseCommand =
                    new(
                        cmdText: databaseCreationSql,
                        connection: masterConnection);

                await createDatabaseCommand.ExecuteNonQueryAsync();
            }

            #endregion


            #region Build Track Geometry Database Connection String

            string databaseConnectionString =
                BuildDatabaseConnectionString(
                    baseConnectionString: connectionString,
                    databaseName: TrackGeometryDatabaseName,
                    pooling: null);

            #endregion


            #region Define Database Table Structure

            const string createTablesSql = """
                SET XACT_ABORT ON;

                BEGIN TRY

                    BEGIN TRANSACTION;
                    /* =============================================================
                       PROJECT
                       Soft deletion:
                           IsDeleted = 0 -> Active
                           IsDeleted = 1 -> Deleted / retired
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.Project', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[Project]
                        (
                            [Project_ID] int IDENTITY(1,1) NOT NULL,
                            [ProjectName] nvarchar(200) NOT NULL,
                            [ProjectStartDate] date NOT NULL
                                CONSTRAINT [DF_Project_ProjectStartDate]
                                DEFAULT (CONVERT(date, GETDATE())),
                            [TimeZoneId] nvarchar(200) NULL,
                            [DefaultReportOutputPath] nvarchar(1000) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_Project_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_Project]
                                PRIMARY KEY CLUSTERED ([Project_ID]),

                            CONSTRAINT [UQ_Project_ProjectName]
                                UNIQUE ([ProjectName])
                        );

                    END;

                    IF COL_LENGTH(N'dbo.Project', N'TimeZoneId') IS NULL
                    BEGIN
                        ALTER TABLE [dbo].[Project]
                            ADD [TimeZoneId] nvarchar(200) NULL;
                    END;

                    IF COL_LENGTH(N'dbo.Project', N'DefaultReportOutputPath') IS NULL
                    BEGIN
                        ALTER TABLE [dbo].[Project]
                            ADD [DefaultReportOutputPath] nvarchar(1000) NULL;
                    END;

                    /* =============================================================
                       DATABASE WRITE HISTORY
                       One row per attempted DBTrackGeometry writing activity.
                       ReportUTCtime is report epoch UTC; WriteUTCtime is actual
                       database-writing UTC. Duplicate report epochs are permitted.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.DatabaseWriteHistory', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[DatabaseWriteHistory]
                        (
                            [WriteHistory_ID] int IDENTITY(1,1) NOT NULL,
                            [Project_ID] int NOT NULL,
                            [ReportUTCtime] datetime2(0) NOT NULL,
                            [WriteUTCtime] datetime2(0) NOT NULL,
                            [Outcome] nvarchar(30) NOT NULL,
                            [Details] nvarchar(max) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_DatabaseWriteHistory_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_DatabaseWriteHistory]
                                PRIMARY KEY CLUSTERED ([WriteHistory_ID]),

                            CONSTRAINT [CK_DatabaseWriteHistory_Outcome]
                                CHECK ([Outcome] IN
                                    (N'Success', N'PartialSuccess', N'Failed')),

                            CONSTRAINT [FK_DatabaseWriteHistory_Project]
                                FOREIGN KEY ([Project_ID])
                                REFERENCES [dbo].[Project] ([Project_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_DatabaseWriteHistory_Project_ReportUTCtime]
                            ON [dbo].[DatabaseWriteHistory]
                            ([Project_ID], [ReportUTCtime]);

                        CREATE INDEX [IX_DatabaseWriteHistory_WriteUTCtime]
                            ON [dbo].[DatabaseWriteHistory] ([WriteUTCtime]);

                    END;

                    /* =============================================================
                       POINT NAME
                       ReplacementName must be populated by the application.
                       If no replacement name exists:
                           ReplacementName = PointName

                       Soft deletion:
                           IsDeleted = 0 -> Active
                           IsDeleted = 1 -> Deleted / retired
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.PointName', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[PointName]
                        (
                            [PointName_ID] int IDENTITY(1,1) NOT NULL,
                            [PointName] nvarchar(50) NOT NULL,
                            [ReplacementName] nvarchar(50) NOT NULL,
                            [Project_ID] int NOT NULL,
                            [LatestReading] datetime2(0) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_PointName_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_PointName]
                                PRIMARY KEY CLUSTERED ([PointName_ID]),

                            CONSTRAINT [UQ_PointName_Project_PointName]
                                UNIQUE ([Project_ID], [PointName]),

                            CONSTRAINT [FK_PointName_Project]
                                FOREIGN KEY ([Project_ID])
                                REFERENCES [dbo].[Project] ([Project_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                    END;

                    /* =============================================================
                       GEOTECHNICAL SENSORS
                       SensorName / ReplacementName identify the sensor within
                       the active project. SensorType is a user descriptor such
                       as Tilt, Vibration, etc.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.GeotecSensors', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[GeotecSensors]
                        (
                            [SensorID] int IDENTITY(1,1) NOT NULL,
                            [SensorName] nvarchar(50) NOT NULL,
                            [ReplacementName] nvarchar(50) NOT NULL,
                            [SensorType] nvarchar(50) NOT NULL,
                            [Project_ID] int NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_GeotecSensors_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_GeotecSensors]
                                PRIMARY KEY CLUSTERED ([SensorID]),

                            CONSTRAINT [UQ_GeotecSensors_Project_SensorName]
                                UNIQUE ([Project_ID], [SensorName]),

                            CONSTRAINT [FK_GeotecSensors_Project]
                                FOREIGN KEY ([Project_ID])
                                REFERENCES [dbo].[Project] ([Project_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_GeotecSensors_Project_SensorType]
                            ON [dbo].[GeotecSensors]
                            ([Project_ID], [SensorType]);

                    END;

                    /* =============================================================
                       COORDINATES REFERENCE
                       One shared reference-coordinate table for both monitoring
                       points and geotechnical sensors. Exactly one owner key must
                       be populated on each row:

                           PointName_ID != NULL, SensorID = NULL
                       OR
                           PointName_ID = NULL, SensorID != NULL

                       Coordinates are stored to six decimal places.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.CoordinatesReference', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[CoordinatesReference]
                        (
                            [Coordinate_ID] int IDENTITY(1,1) NOT NULL,
                            [PointName_ID] int NULL,
                            [SensorID] int NULL,
                            [Eref] decimal(18,6) NOT NULL,
                            [Nref] decimal(18,6) NOT NULL,
                            [Href] decimal(18,6) NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_CoordinatesReference_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_CoordinatesReference]
                                PRIMARY KEY CLUSTERED ([Coordinate_ID]),

                            CONSTRAINT [CK_CoordinatesReference_OneOwner]
                                CHECK
                                (
                                    ([PointName_ID] IS NOT NULL AND [SensorID] IS NULL)
                                    OR
                                    ([PointName_ID] IS NULL AND [SensorID] IS NOT NULL)
                                ),

                            CONSTRAINT [FK_CoordinatesReference_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION,

                            CONSTRAINT [FK_CoordinatesReference_GeotecSensors]
                                FOREIGN KEY ([SensorID])
                                REFERENCES [dbo].[GeotecSensors] ([SensorID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE UNIQUE INDEX
                            [UX_CoordinatesReference_Active_PointName]
                            ON [dbo].[CoordinatesReference] ([PointName_ID])
                            WHERE
                                [IsDeleted] = 0
                                AND [PointName_ID] IS NOT NULL;

                        CREATE UNIQUE INDEX
                            [UX_CoordinatesReference_Active_SensorID]
                            ON [dbo].[CoordinatesReference] ([SensorID])
                            WHERE
                                [IsDeleted] = 0
                                AND [SensorID] IS NOT NULL;

                    END;

                    /* =============================================================
                       TOP-OF-RAIL REFERENCE
                       Reference ToR from Survey worksheet column G.
                       One active row per point. Value stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ToRReference', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ToRReference]
                        (
                            [ToRReference_ID] int IDENTITY(1,1) NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [ToR] decimal(18,6) NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_ToRReference_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_ToRReference]
                                PRIMARY KEY CLUSTERED ([ToRReference_ID]),

                            CONSTRAINT [FK_ToRReference_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE UNIQUE INDEX
                            [UX_ToRReference_Active_PointName_ID]
                            ON [dbo].[ToRReference] ([PointName_ID])
                            WHERE [IsDeleted] = 0;

                    END;

                    /* =============================================================
                       COORDINATES CURRENT
                       Time-stamped coordinate epoch history for monitoring points.
                       Coordinates stored in metres. Missing values are SQL NULL.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.CoordinatesEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[CoordinatesEpochs]
                        (
                            [PointName_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [LatestReading] datetime2(0) NULL,
                            [ReadingCount] int NULL,
                            [E] decimal(18,4) NULL,
                            [N] decimal(18,4) NULL,
                            [H] decimal(18,4) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_CoordinatesEpochs_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_CoordinatesEpochs]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [CK_CoordinatesEpochs_ReadingCount]
                                CHECK ([ReadingCount] IS NULL OR [ReadingCount] >= 0),

                            CONSTRAINT [FK_CoordinatesEpochs_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_CoordinatesEpochs_UTCtime]
                            ON [dbo].[CoordinatesEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       TOP-OF-RAIL CURRENT
                       Time-stamped ToR history read from recalculated Reference
                       worksheet column P. Do not recalculate as H + ToRoffset.
                       Missing values are SQL NULL. Value stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ToREpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ToREpochs]
                        (
                            [PointName_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [ToR] decimal(18,4) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_ToREpochs_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_ToREpochs]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_ToREpochs_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_ToREpochs_UTCtime]
                            ON [dbo].[ToREpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       TRACK
                       Each Track belongs to one Project.
                       TrackName is unique within that Project.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.Track', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[Track]
                        (
                            [Track_ID] int IDENTITY(1,1) NOT NULL,
                            [Project_ID] int NOT NULL,
                            [TrackName] nvarchar(200) NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_Track_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_Track]
                                PRIMARY KEY CLUSTERED ([Track_ID]),

                            CONSTRAINT [UQ_Track_Project_TrackName]
                                UNIQUE ([Project_ID], [TrackName]),

                            CONSTRAINT [FK_Track_Project]
                                FOREIGN KEY ([Project_ID])
                                REFERENCES [dbo].[Project] ([Project_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                    END;

                    /* =============================================================
                       PRISM PAIRS
                       Each pair belongs to one Track and has an explicit order.
                       Left and Right points must be different.
                       Same-project validation is performed by the application.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.PrismPairs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[PrismPairs]
                        (
                            [PrismPair_ID] int IDENTITY(1,1) NOT NULL,
                            [Track_ID] int NOT NULL,
                            [PairOrder] int NOT NULL,
                            [Left_ID] int NOT NULL,
                            [Right_ID] int NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_PrismPairs_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_PrismPairs]
                                PRIMARY KEY CLUSTERED ([PrismPair_ID]),

                            CONSTRAINT [UQ_PrismPairs_Left_Right]
                                UNIQUE ([Left_ID], [Right_ID]),

                            CONSTRAINT [UQ_PrismPairs_Track_PairOrder]
                                UNIQUE ([Track_ID], [PairOrder]),

                            CONSTRAINT [CK_PrismPairs_DifferentPoints]
                                CHECK ([Left_ID] <> [Right_ID]),

                            CONSTRAINT [CK_PrismPairs_PairOrder]
                                CHECK ([PairOrder] > 0),

                            CONSTRAINT [FK_PrismPairs_Track]
                                FOREIGN KEY ([Track_ID])
                                REFERENCES [dbo].[Track] ([Track_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION,

                            CONSTRAINT [FK_PrismPairs_LeftPoint]
                                FOREIGN KEY ([Left_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION,

                            CONSTRAINT [FK_PrismPairs_RightPoint]
                                FOREIGN KEY ([Right_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_PrismPairs_Track_ID]
                            ON [dbo].[PrismPairs] ([Track_ID]);

                    END;

                                        /* =============================================================
                       PRISM ARRAY TYPE
                       Stable reference-data lookup for permitted PrismArray types.

                       ArrayType_ID values are permanent programmatic identities.
                       This table is not soft-deleted and is not user-managed.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.PrismArrayType', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[PrismArrayType]
                        (
                            [ArrayType_ID] int NOT NULL,
                            [ArrayTypeName] nvarchar(100) NOT NULL,

                            CONSTRAINT [PK_PrismArrayType]
                                PRIMARY KEY CLUSTERED ([ArrayType_ID]),

                            CONSTRAINT [UQ_PrismArrayType_ArrayTypeName]
                                UNIQUE ([ArrayTypeName])
                        );

                        INSERT INTO [dbo].[PrismArrayType]
                        (
                            [ArrayType_ID],
                            [ArrayTypeName]
                        )
                        VALUES
                            (1, N'Structural Array'),
                            (2, N'Tunnel Convergence'),
                            (3, N'PrismCrackGauge'),
                            (4, N'PrismTilt (MperM)');

                    END;

                    /* =============================================================
                       PRISM ARRAY
                       Common array definition.

                       ArrayType references dbo.PrismArrayType and must therefore
                       contain a defined database lookup identity.

                       Array names are unique only among ACTIVE arrays within a
                       project. A soft-deleted array name may be reused.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.PrismArray', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[PrismArray]
                        (
                            [Array_ID] int IDENTITY(1,1) NOT NULL,
                            [Project_ID] int NOT NULL,
                            [ArrayName] nvarchar(200) NOT NULL,
                            [ArrayType] int NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_PrismArray_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_PrismArray]
                                PRIMARY KEY CLUSTERED ([Array_ID]),

                            CONSTRAINT [FK_PrismArray_Project]
                                FOREIGN KEY ([Project_ID])
                                REFERENCES [dbo].[Project] ([Project_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION,

                            CONSTRAINT [FK_PrismArray_PrismArrayType]
                                FOREIGN KEY ([ArrayType])
                                REFERENCES [dbo].[PrismArrayType] ([ArrayType_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_PrismArray_Project_ID]
                            ON [dbo].[PrismArray] ([Project_ID]);

                        CREATE UNIQUE INDEX
                            [UX_PrismArray_Active_Project_ArrayName]
                            ON [dbo].[PrismArray]
                            (
                                [Project_ID],
                                [ArrayName]
                            )
                            WHERE [IsDeleted] = 0;

                    END;

                                        /* =============================================================
                       PRISM ARRAY POINT
                       Common A-E point membership for every array type.

                       A point may belong to one ACTIVE array of each type.
                       Each array may contain one point for each role A-E.

                       Soft-deleted membership rows are retained so historical
                       array data remain referentially intact.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.PrismArrayPoint', N'U') IS NULL
                    BEGIN
                        CREATE TABLE [dbo].[PrismArrayPoint]
                        (
                            [Array_ID] int NOT NULL,
                            [PointRole] char(1) NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_PrismArrayPoint_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_PrismArrayPoint]
                                PRIMARY KEY CLUSTERED
                                ([Array_ID], [PointRole]),

                            CONSTRAINT [UQ_PrismArrayPoint_Array_Point]
                                UNIQUE ([Array_ID], [PointName_ID]),

                            CONSTRAINT [CK_PrismArrayPoint_PointRole]
                                CHECK ([PointRole] IN ('A','B','C','D','E')),

                            CONSTRAINT [FK_PrismArrayPoint_Array]
                                FOREIGN KEY ([Array_ID])
                                REFERENCES [dbo].[PrismArray] ([Array_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION,

                            CONSTRAINT [FK_PrismArrayPoint_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE NONCLUSTERED INDEX
                            [IX_PrismArrayPoint_Active_PointName_ID]
                            ON [dbo].[PrismArrayPoint]
                            (
                                [PointName_ID]
                            )
                            INCLUDE
                            (
                                [Array_ID],
                                [PointRole]
                            )
                            WHERE [IsDeleted] = 0;
                    END;

                    /* =============================================================
                       dH EPOCHS
                       Individual epoch values stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.DhEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[DhEpochs]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [dH] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_DhEpochs_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_DhEpochs]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_DhEpochs_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_DhEpochs_UTCtime]
                            ON [dbo].[DhEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       SLEW EPOCHS
                       Individual epoch values stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.SlewEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[SlewEpochs]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [Slew] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_SlewEpochs_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_SlewEpochs]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_SlewEpochs_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_SlewEpochs_UTCtime]
                            ON [dbo].[SlewEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       TOP EPOCHS
                       Individual epoch values stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.TopEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[TopEpochs]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [Top] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_TopEpochs_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_TopEpochs]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_TopEpochs_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_TopEpochs_UTCtime]
                            ON [dbo].[TopEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       VERSINE EPOCHS
                       Stores the current absolute signed horizontal Versine
                       independently for each Left and Right rail point.

                       RightVersine -> Right rail PointName_ID
                       LeftVersine  -> Left rail PointName_ID

                       Stored value:
                           signed horizontal Versine in metres

                       Positive = current rail point lies left of chord A-C when
                                  looking in the direction of increasing mileage.
                       Negative = current rail point lies right of chord A-C.

                       First/last point in a rail section, missing-coordinate
                       conditions, invalid A-C chord geometry and other unavailable
                       measurements are stored as SQL NULL.

                       Epoch identity:
                           (PointName_ID, UTCtime)
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.VersineEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[VersineEpochs]
                        (
                            [PointName_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [Versine] decimal(18,4) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_VersineEpochs_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_VersineEpochs]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_VersineEpochs_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_VersineEpochs_UTCtime]
                            ON [dbo].[VersineEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       CANT EPOCHS
                       Individual epoch values stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.CantEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[CantEpochs]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [Cant] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_CantEpochs_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_CantEpochs]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCtime]),

                            CONSTRAINT [FK_CantEpochs_PrismPairs]
                                FOREIGN KEY ([PrismPair_ID])
                                REFERENCES [dbo].[PrismPairs] ([PrismPair_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_CantEpochs_UTCtime]
                            ON [dbo].[CantEpochs] ([UTCtime]);

                    END;

                                        /* =============================================================
                       SHORT TWIST EPOCHS
                       Stores both Track Geometry Short Twist representations:

                           ShortTwist
                               Signed Cant difference in metres over the 3 m
                               pair interval.

                           ShortTwistRatio
                               Absolute workbook Short Twist ratio rounded to an
                               integer.

                       ShortTwist = CurrentPair.Cant - PreviousPair.Cant

                       ShortTwistRatio = ABS(3.0 / ShortTwistMetres)

                       If ShortTwistMetres is exactly zero, 10000000 is the valid
                       workbook-compatibility ratio value.

                       Missing or unavailable values are stored as SQL NULL.
                       The ratio must not be reconstructed from the rounded
                       ShortTwist database value.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ShortTwistEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ShortTwistEpochs]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [ShortTwist] decimal(18,4) NULL,
                            [ShortTwistRatio] int NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_ShortTwistEpochs_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_ShortTwistEpochs]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCtime]),

                            CONSTRAINT [CK_ShortTwistEpochs_ShortTwistRatio]
                                CHECK
                                (
                                    [ShortTwistRatio] IS NULL
                                    OR [ShortTwistRatio] >= 0
                                ),

                            CONSTRAINT [FK_ShortTwistEpochs_PrismPairs]
                                FOREIGN KEY ([PrismPair_ID])
                                REFERENCES [dbo].[PrismPairs] ([PrismPair_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_ShortTwistEpochs_UTCtime]
                            ON [dbo].[ShortTwistEpochs] ([UTCtime]);

                    END;

                                        /* =============================================================
                       LONG TWIST EPOCHS
                       Stores both Track Geometry Long Twist representations:

                           LongTwist
                               Signed Cant difference in metres over the 15 m
                               interval.

                           LongTwistRatio
                               Absolute workbook Long Twist ratio rounded to an
                               integer.

                       LongTwist = CurrentPair.Cant - CantFivePairsEarlier

                       LongTwistRatio = ABS(15.0 / LongTwistMetres)

                       If LongTwistMetres is exactly zero, 10000000 is the valid
                       workbook-compatibility ratio value.

                       Missing or unavailable values are stored as SQL NULL.
                       The ratio must not be reconstructed from the rounded
                       LongTwist database value.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.LongTwistEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[LongTwistEpochs]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [LongTwist] decimal(18,4) NULL,
                            [LongTwistRatio] int NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_LongTwistEpochs_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_LongTwistEpochs]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCtime]),

                            CONSTRAINT [CK_LongTwistEpochs_LongTwistRatio]
                                CHECK
                                (
                                    [LongTwistRatio] IS NULL
                                    OR [LongTwistRatio] >= 0
                                ),

                            CONSTRAINT [FK_LongTwistEpochs_PrismPairs]
                                FOREIGN KEY ([PrismPair_ID])
                                REFERENCES [dbo].[PrismPairs] ([PrismPair_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_LongTwistEpochs_UTCtime]
                            ON [dbo].[LongTwistEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       COORDINATES DAILY
                       UTC calendar-day arithmetic mean of non-NULL epoch values.
                       Daily row timestamp is 12:00:00 UTC.
                       If all values are NULL, retain a row with NULL values.
                       Coordinates stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.CoordinatesDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[CoordinatesDaily]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [E] decimal(18,4) NULL,
                            [N] decimal(18,4) NULL,
                            [H] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_CoordinatesDaily_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_CoordinatesDaily]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_CoordinatesDaily_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_CoordinatesDaily_UTCtime]
                            ON [dbo].[CoordinatesDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       TOP-OF-RAIL DAILY
                       UTC calendar-day arithmetic mean of non-NULL ToREpochs.
                       Daily row timestamp is 12:00:00 UTC. If all epoch values are
                       NULL, retain a Daily row with ToR = NULL.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ToRDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ToRDaily]
                        (
                            [PointName_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [ToR] decimal(18,4) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_ToRDaily_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_ToRDaily]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_ToRDaily_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_ToRDaily_UTCtime]
                            ON [dbo].[ToRDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       dH DAILY
                       UTC calendar-day arithmetic mean of non-NULL DhEpochs.
                       Daily row timestamp is 12:00:00 UTC.
                       Value stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.DhDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[DhDaily]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [dH] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_DhDaily_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_DhDaily]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_DhDaily_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_DhDaily_UTCtime]
                            ON [dbo].[DhDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       SLEW DAILY
                       UTC calendar-day arithmetic mean of non-NULL SlewEpochs.
                       Daily row timestamp is 12:00:00 UTC.
                       Value stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.SlewDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[SlewDaily]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [Slew] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_SlewDaily_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_SlewDaily]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_SlewDaily_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_SlewDaily_UTCtime]
                            ON [dbo].[SlewDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       TOP DAILY
                       UTC calendar-day arithmetic mean of non-NULL TopEpochs.
                       Daily row timestamp is 12:00:00 UTC.
                       Value stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.TopDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[TopDaily]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [Top] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_TopDaily_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_TopDaily]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_TopDaily_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_TopDaily_UTCtime]
                            ON [dbo].[TopDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       VERSINE DAILY
                       UTC calendar-day arithmetic mean of non-NULL signed
                       VersineEpochs.Versine values for each rail point.

                       Sign is preserved during aggregation.
                       Daily row timestamp is 12:00:00 UTC for the applicable date.

                       If all applicable epoch values are NULL, retain a Daily row
                       with Versine = NULL.

                       Stored value:
                           signed horizontal Versine in metres

                       Daily identity:
                           (PointName_ID, UTCtime)
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.VersineDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[VersineDaily]
                        (
                            [PointName_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [Versine] decimal(18,4) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_VersineDaily_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_VersineDaily]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_VersineDaily_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_VersineDaily_UTCtime]
                            ON [dbo].[VersineDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       CANT DAILY
                       UTC calendar-day arithmetic mean of non-NULL CantEpochs.
                       Daily row timestamp is 12:00:00 UTC.
                       Value stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.CantDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[CantDaily]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [Cant] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_CantDaily_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_CantDaily]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCtime]),

                            CONSTRAINT [FK_CantDaily_PrismPairs]
                                FOREIGN KEY ([PrismPair_ID])
                                REFERENCES [dbo].[PrismPairs] ([PrismPair_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_CantDaily_UTCtime]
                            ON [dbo].[CantDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       SHORT TWIST DAILY
                       Daily mean of signed ShortTwistEpochs Cant differences.
                       Value stored in metres over the 3 m pair interval.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ShortTwistDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ShortTwistDaily]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [ShortTwist] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_ShortTwistDaily_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_ShortTwistDaily]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCtime]),

                            CONSTRAINT [FK_ShortTwistDaily_PrismPairs]
                                FOREIGN KEY ([PrismPair_ID])
                                REFERENCES [dbo].[PrismPairs] ([PrismPair_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_ShortTwistDaily_UTCtime]
                            ON [dbo].[ShortTwistDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       LONG TWIST DAILY
                       Daily mean of signed LongTwistEpochs Cant differences.
                       Value stored in metres over the 15 m interval.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.LongTwistDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[LongTwistDaily]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [LongTwist] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_LongTwistDaily_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_LongTwistDaily]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCtime]),

                            CONSTRAINT [FK_LongTwistDaily_PrismPairs]
                                FOREIGN KEY ([PrismPair_ID])
                                REFERENCES [dbo].[PrismPairs] ([PrismPair_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_LongTwistDaily_UTCtime]
                            ON [dbo].[LongTwistDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       STRUCTURAL ARRAY EPOCHS
                       Point displacement relative to reference coordinates.
                       UTCtime is UTC date/time rounded to the nearest second.
                       Values stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.StructuralArrayEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[StructuralArrayEpochs]
                        (
                            [Array_ID] int NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [dE] decimal(18,4) NULL,
                            [dN] decimal(18,4) NULL,
                            [dH] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_StructuralArrayEpochs_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_StructuralArrayEpochs]
                                PRIMARY KEY CLUSTERED
                                ([Array_ID], [PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_StructuralArrayEpochs_ArrayPoint]
                                FOREIGN KEY ([Array_ID], [PointName_ID])
                                REFERENCES [dbo].[PrismArrayPoint]
                                ([Array_ID], [PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_StructuralArrayEpochs_UTCtime]
                            ON [dbo].[StructuralArrayEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       STRUCTURAL ARRAY DAILY
                       Daily derived point displacement dataset.
                       UTCtime remains a full UTC date/time to the nearest second.
                       Values stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.StructuralArrayDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[StructuralArrayDaily]
                        (
                            [Array_ID] int NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [dE] decimal(18,4) NULL,
                            [dN] decimal(18,4) NULL,
                            [dH] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_StructuralArrayDaily_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_StructuralArrayDaily]
                                PRIMARY KEY CLUSTERED
                                ([Array_ID], [PointName_ID], [UTCtime]),

                            CONSTRAINT [FK_StructuralArrayDaily_ArrayPoint]
                                FOREIGN KEY ([Array_ID], [PointName_ID])
                                REFERENCES [dbo].[PrismArrayPoint]
                                ([Array_ID], [PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_StructuralArrayDaily_UTCtime]
                            ON [dbo].[StructuralArrayDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       TUNNEL CONVERGENCE EPOCHS
                       Signed change in 3D chord length relative to reference.

                       dChord = current 3D chord length - reference 3D chord length

                       Complete chord set:
                           AB AC AD AE BC BD BE CD CE DE

                       UTCtime is UTC date/time rounded to the nearest second.
                       Values stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.TunnelConvergenceEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[TunnelConvergenceEpochs]
                        (
                            [Array_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [dAB] decimal(18,4) NULL,
                            [dAC] decimal(18,4) NULL,
                            [dAD] decimal(18,4) NULL,
                            [dAE] decimal(18,4) NULL,
                            [dBC] decimal(18,4) NULL,
                            [dBD] decimal(18,4) NULL,
                            [dBE] decimal(18,4) NULL,
                            [dCD] decimal(18,4) NULL,
                            [dCE] decimal(18,4) NULL,
                            [dDE] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_TunnelConvergenceEpochs_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_TunnelConvergenceEpochs]
                                PRIMARY KEY CLUSTERED
                                ([Array_ID], [UTCtime]),

                            CONSTRAINT [FK_TunnelConvergenceEpochs_Array]
                                FOREIGN KEY ([Array_ID])
                                REFERENCES [dbo].[PrismArray] ([Array_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_TunnelConvergenceEpochs_UTCtime]
                            ON [dbo].[TunnelConvergenceEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       TUNNEL CONVERGENCE DAILY
                       Daily derived signed change in 3D chord length.
                       UTCtime remains a full UTC date/time to the nearest second.
                       Values stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.TunnelConvergenceDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[TunnelConvergenceDaily]
                        (
                            [Array_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [dAB] decimal(18,4) NULL,
                            [dAC] decimal(18,4) NULL,
                            [dAD] decimal(18,4) NULL,
                            [dAE] decimal(18,4) NULL,
                            [dBC] decimal(18,4) NULL,
                            [dBD] decimal(18,4) NULL,
                            [dBE] decimal(18,4) NULL,
                            [dCD] decimal(18,4) NULL,
                            [dCE] decimal(18,4) NULL,
                            [dDE] decimal(18,4) NULL,

                            [IsDeleted] bit NOT NULL

                                CONSTRAINT [DF_TunnelConvergenceDaily_IsDeleted]

                                DEFAULT (0),


                            CONSTRAINT [PK_TunnelConvergenceDaily]
                                PRIMARY KEY CLUSTERED
                                ([Array_ID], [UTCtime]),

                            CONSTRAINT [FK_TunnelConvergenceDaily_Array]
                                FOREIGN KEY ([Array_ID])
                                REFERENCES [dbo].[PrismArray] ([Array_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_TunnelConvergenceDaily_UTCtime]
                            ON [dbo].[TunnelConvergenceDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       PRISM CRACK GAUGE EPOCHS
                       Derived A-B crack-gauge geometry for each report epoch.

                       All measurements are stored in metres:
                           d2D = Current 2D A-B distance - Reference 2D A-B distance
                           d3D = Current 3D A-B distance - Reference 3D A-B distance
                           dH  = Current (B.H - A.H) - Reference (B.Href - A.Href)

                       Measurements are calculated by TrackGeometryReport and are
                       written to this table by the database export workflow.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.PrismCrackGaugeEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[PrismCrackGaugeEpochs]
                        (
                            [Array_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [d2D] decimal(18,4) NULL,
                            [d3D] decimal(18,4) NULL,
                            [dH] decimal(18,4) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_PrismCrackGaugeEpochs_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_PrismCrackGaugeEpochs]
                                PRIMARY KEY CLUSTERED
                                ([Array_ID], [UTCtime]),

                            CONSTRAINT [FK_PrismCrackGaugeEpochs_PrismArray]
                                FOREIGN KEY ([Array_ID])
                                REFERENCES [dbo].[PrismArray] ([Array_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_PrismCrackGaugeEpochs_UTCtime]
                            ON [dbo].[PrismCrackGaugeEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       PRISM CRACK GAUGE DAILY
                       UTC calendar-day summary of PrismCrackGaugeEpochs data.

                       All measurements are stored in metres. Daily values are
                       populated by the Track Geometry database-writing workflow.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.PrismCrackGaugeDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[PrismCrackGaugeDaily]
                        (
                            [Array_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [d2D] decimal(18,4) NULL,
                            [d3D] decimal(18,4) NULL,
                            [dH] decimal(18,4) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_PrismCrackGaugeDaily_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_PrismCrackGaugeDaily]
                                PRIMARY KEY CLUSTERED
                                ([Array_ID], [UTCtime]),

                            CONSTRAINT [FK_PrismCrackGaugeDaily_PrismArray]
                                FOREIGN KEY ([Array_ID])
                                REFERENCES [dbo].[PrismArray] ([Array_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_PrismCrackGaugeDaily_UTCtime]
                            ON [dbo].[PrismCrackGaugeDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       PRISM TILT EPOCHS
                       Two-point PrismTilt array results generated by
                       TrackGeometryReport.

                       TiltX_MperM and TiltY_MperM are stored in metres/metre
                       (m/m) to six decimal places.

                       Array_ID references dbo.PrismArray. PrismTilt arrays use
                       point roles A and B in the Arrays configuration workflow.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.PrismTiltEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[PrismTiltEpochs]
                        (
                            [Array_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [TiltX_MperM] decimal(18,6) NULL,
                            [TiltY_MperM] decimal(18,6) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_PrismTiltEpochs_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_PrismTiltEpochs]
                                PRIMARY KEY CLUSTERED
                                ([Array_ID], [UTCtime]),

                            CONSTRAINT [FK_PrismTiltEpochs_PrismArray]
                                FOREIGN KEY ([Array_ID])
                                REFERENCES [dbo].[PrismArray] ([Array_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_PrismTiltEpochs_UTCtime]
                            ON [dbo].[PrismTiltEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       PRISM TILT DAILY
                       Daily PrismTilt values generated by TrackGeometryReport.

                       TiltX_MperM and TiltY_MperM are stored in metres/metre
                       (m/m) to six decimal places.

                       UTCtime follows the existing daily-table convention.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.PrismTiltDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[PrismTiltDaily]
                        (
                            [Array_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [TiltX_MperM] decimal(18,6) NULL,
                            [TiltY_MperM] decimal(18,6) NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_PrismTiltDaily_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_PrismTiltDaily]
                                PRIMARY KEY CLUSTERED
                                ([Array_ID], [UTCtime]),

                            CONSTRAINT [FK_PrismTiltDaily_PrismArray]
                                FOREIGN KEY ([Array_ID])
                                REFERENCES [dbo].[PrismArray] ([Array_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_PrismTiltDaily_UTCtime]
                            ON [dbo].[PrismTiltDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       TILT EPOCHS
                       All Tilt values are stored to six decimal places.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.TiltEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[TiltEpochs]
                        (
                            [SensorID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [TiltA] decimal(18,6) NOT NULL,
                            [TiltB] decimal(18,6) NOT NULL,
                            [TiltC] decimal(18,6) NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_TiltEpochs_IsDeleted]
                                DEFAULT (0),
                            [ReplacementName] nvarchar(50) NOT NULL,

                            CONSTRAINT [PK_TiltEpochs]
                                PRIMARY KEY CLUSTERED
                                ([SensorID], [UTCtime]),

                            CONSTRAINT [FK_TiltEpochs_GeotecSensors]
                                FOREIGN KEY ([SensorID])
                                REFERENCES [dbo].[GeotecSensors] ([SensorID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_TiltEpochs_UTCtime]
                            ON [dbo].[TiltEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       TILT DAILY
                       All Tilt values are stored to six decimal places.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.TiltDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[TiltDaily]
                        (
                            [SensorID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [TiltA] decimal(18,6) NOT NULL,
                            [TiltB] decimal(18,6) NOT NULL,
                            [TiltC] decimal(18,6) NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_TiltDaily_IsDeleted]
                                DEFAULT (0),
                            [ReplacementName] nvarchar(50) NOT NULL,

                            CONSTRAINT [PK_TiltDaily]
                                PRIMARY KEY CLUSTERED
                                ([SensorID], [UTCtime]),

                            CONSTRAINT [FK_TiltDaily_GeotecSensors]
                                FOREIGN KEY ([SensorID])
                                REFERENCES [dbo].[GeotecSensors] ([SensorID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_TiltDaily_UTCtime]
                            ON [dbo].[TiltDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       VIBRATION EPOCHS
                       All Vibration values are stored to six decimal places.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.VibrationEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[VibrationEpochs]
                        (
                            [SensorID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [Vibration] decimal(18,6) NOT NULL,
                            [PPV] decimal(18,6) NOT NULL,
                            [PPA] decimal(18,6) NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_VibrationEpochs_IsDeleted]
                                DEFAULT (0),
                            [ReplacementName] nvarchar(50) NOT NULL,

                            CONSTRAINT [PK_VibrationEpochs]
                                PRIMARY KEY CLUSTERED
                                ([SensorID], [UTCtime]),

                            CONSTRAINT [FK_VibrationEpochs_GeotecSensors]
                                FOREIGN KEY ([SensorID])
                                REFERENCES [dbo].[GeotecSensors] ([SensorID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_VibrationEpochs_UTCtime]
                            ON [dbo].[VibrationEpochs] ([UTCtime]);

                    END;

                    /* =============================================================
                       VIBRATION HOURLY
                       All Vibration values are stored to six decimal places.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.VibrationHourly', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[VibrationHourly]
                        (
                            [SensorID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [Vibration] decimal(18,6) NOT NULL,
                            [PPV] decimal(18,6) NOT NULL,
                            [PPA] decimal(18,6) NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_VibrationHourly_IsDeleted]
                                DEFAULT (0),
                            [ReplacementName] nvarchar(50) NOT NULL,

                            CONSTRAINT [PK_VibrationHourly]
                                PRIMARY KEY CLUSTERED
                                ([SensorID], [UTCtime]),

                            CONSTRAINT [FK_VibrationHourly_GeotecSensors]
                                FOREIGN KEY ([SensorID])
                                REFERENCES [dbo].[GeotecSensors] ([SensorID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_VibrationHourly_UTCtime]
                            ON [dbo].[VibrationHourly] ([UTCtime]);

                    END;

                    /* =============================================================
                       VIBRATION DAILY
                       All Vibration values are stored to six decimal places.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.VibrationDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[VibrationDaily]
                        (
                            [SensorID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [Vibration] decimal(18,6) NOT NULL,
                            [PPV] decimal(18,6) NOT NULL,
                            [PPA] decimal(18,6) NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_VibrationDaily_IsDeleted]
                                DEFAULT (0),
                            [ReplacementName] nvarchar(50) NOT NULL,

                            CONSTRAINT [PK_VibrationDaily]
                                PRIMARY KEY CLUSTERED
                                ([SensorID], [UTCtime]),

                            CONSTRAINT [FK_VibrationDaily_GeotecSensors]
                                FOREIGN KEY ([SensorID])
                                REFERENCES [dbo].[GeotecSensors] ([SensorID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_VibrationDaily_UTCtime]
                            ON [dbo].[VibrationDaily] ([UTCtime]);

                    END;

                    /* =============================================================
                       CHART TYPE
                       Reusable catalogue entry defining the DB source and
                       rendering behaviour for one logical chart type.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ChartType', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ChartType]
                        (
                            [ChartType_ID] int IDENTITY(1,1) NOT NULL,
                            [ChartTypeKey] nvarchar(100) NOT NULL,
                            [DisplayName] nvarchar(200) NOT NULL,
                            [DataSourceTable] sysname NOT NULL,
                            [TimestampColumnName] sysname NOT NULL,
                            [EntityIdColumnName] sysname NOT NULL,
                            [EntityKind] nvarchar(30) NOT NULL,
                            [DefaultTitleTemplate] nvarchar(500) NOT NULL,
                            [DefaultYAxisTitle] nvarchar(200) NOT NULL,
                            [DefaultXAxisTitle] nvarchar(200) NOT NULL
                                CONSTRAINT [DF_ChartType_DefaultXAxisTitle]
                                DEFAULT (N'Date / Time'),
                            [SupportsReferenceAdjustment] bit NOT NULL
                                CONSTRAINT [DF_ChartType_SupportsReferenceAdjustment]
                                DEFAULT (1),
                            [SupportsTriggerBands] bit NOT NULL
                                CONSTRAINT [DF_ChartType_SupportsTriggerBands]
                                DEFAULT (1),
                            [DefaultTriggerBandsSymmetric] bit NOT NULL
                                CONSTRAINT [DF_ChartType_DefaultTriggerBandsSymmetric]
                                DEFAULT (1),
                            [IsEnabled] bit NOT NULL
                                CONSTRAINT [DF_ChartType_IsEnabled]
                                DEFAULT (1),
                            [DisplayOrder] int NOT NULL
                                CONSTRAINT [DF_ChartType_DisplayOrder]
                                DEFAULT (0),

                            CONSTRAINT [PK_ChartType]
                                PRIMARY KEY CLUSTERED ([ChartType_ID]),

                            CONSTRAINT [UQ_ChartType_ChartTypeKey]
                                UNIQUE ([ChartTypeKey]),

                            CONSTRAINT [CK_ChartType_EntityKind]
                                CHECK
                                (
                                    [EntityKind] IN
                                    (
                                        N'Point',
                                        N'PrismPair',
                                        N'Track',
                                        N'PrismArray',
                                        N'Sensor'
                                    )
                                )
                        );

                    END;

                    /* =============================================================
                       CHART TYPE DATA ELEMENT
                       Defines one selectable data element/value column for a
                       chart type and its display-unit conversion.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ChartTypeDataElement', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ChartTypeDataElement]
                        (
                            [ChartTypeDataElement_ID] int IDENTITY(1,1) NOT NULL,
                            [ChartType_ID] int NOT NULL,
                            [DataElementKey] nvarchar(100) NOT NULL,
                            [DisplayName] nvarchar(200) NOT NULL,
                            [ValueColumnName] sysname NOT NULL,
                            [DisplayUnit] nvarchar(50) NOT NULL,
                            [DisplayScaleFactor] decimal(18,9) NOT NULL
                                CONSTRAINT [DF_ChartTypeDataElement_DisplayScaleFactor]
                                DEFAULT (1.0),
                            [DisplayOrder] int NOT NULL
                                CONSTRAINT [DF_ChartTypeDataElement_DisplayOrder]
                                DEFAULT (0),
                            [IsEnabled] bit NOT NULL
                                CONSTRAINT [DF_ChartTypeDataElement_IsEnabled]
                                DEFAULT (1),

                            CONSTRAINT [PK_ChartTypeDataElement]
                                PRIMARY KEY CLUSTERED ([ChartTypeDataElement_ID]),

                            CONSTRAINT [FK_ChartTypeDataElement_ChartType]
                                FOREIGN KEY ([ChartType_ID])
                                REFERENCES [dbo].[ChartType] ([ChartType_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION,

                            CONSTRAINT [UQ_ChartTypeDataElement]
                                UNIQUE ([ChartType_ID], [DataElementKey])
                        );

                    END;

                    /* =============================================================
                       CHART DEFINITION
                       One committed chart instance for one project.

                       ChartNumber is sequential per project. Soft-deleted chart
                       numbers remain consumed and are not recycled.

                       Only chart configuration is persisted. Generated PNG files
                       are temporary working files and are not stored here.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ChartDefinition', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ChartDefinition]
                        (
                            [ChartDefinition_ID] int IDENTITY(1,1) NOT NULL,
                            [ChartDefinitionGuid] uniqueidentifier NOT NULL
                                CONSTRAINT [DF_ChartDefinition_Guid]
                                DEFAULT (NEWID()),

                            [Project_ID] int NOT NULL,
                            [ChartNumber] int NOT NULL,
                            [ChartName] nvarchar(200) NOT NULL,
                            [ChartType_ID] int NOT NULL,

                            [AutoTitleTemplate] nvarchar(500) NOT NULL,
                            [TitleOverride] nvarchar(500) NULL,

                            [YAxisTitle] nvarchar(200) NOT NULL,
                            [YAxisUnit] nvarchar(50) NOT NULL,
                            [XAxisTitle] nvarchar(200) NOT NULL
                                CONSTRAINT [DF_ChartDefinition_XAxisTitle]
                                DEFAULT (N'Date / Time'),
                            [DisplayTimeZoneId] nvarchar(200) NOT NULL,

                            [UseAutomaticYAxis] bit NOT NULL
                                CONSTRAINT [DF_ChartDefinition_UseAutomaticYAxis]
                                DEFAULT (1),
                            [FixedYAxisMinimum] decimal(18,6) NULL,
                            [FixedYAxisMaximum] decimal(18,6) NULL,

                            [ReferenceLineValue] decimal(18,6) NOT NULL
                                CONSTRAINT [DF_ChartDefinition_ReferenceLineValue]
                                DEFAULT (0),

                            [ShowLegend] bit NOT NULL
                                CONSTRAINT [DF_ChartDefinition_ShowLegend]
                                DEFAULT (1),
                            [LegendPosition] nvarchar(50) NOT NULL
                                CONSTRAINT [DF_ChartDefinition_LegendPosition]
                                DEFAULT (N'LowerCenter'),

                            [PngWidthPixels] int NOT NULL
                                CONSTRAINT [DF_ChartDefinition_PngWidth]
                                DEFAULT (1800),
                            [PngHeightPixels] int NOT NULL
                                CONSTRAINT [DF_ChartDefinition_PngHeight]
                                DEFAULT (600),

                            [ChartTimeWindowMode] nvarchar(20) NOT NULL,
                            [AbsoluteStartUtc] datetime2(3) NULL,
                            [AbsoluteEndUtc] datetime2(3) NULL,
                            [RelativeAnchor] nvarchar(20) NULL,
                            [RelativeStartOffsetSec] bigint NULL,
                            [RelativeEndOffsetSec] bigint NULL,

                            [ReferenceMode] nvarchar(30) NOT NULL
                                CONSTRAINT [DF_ChartDefinition_ReferenceMode]
                                DEFAULT (N'None'),
                            [ReferenceDateTimeUtc] datetime2(3) NULL,
                            [ReferenceBlockSeconds] bigint NULL,
                            [UseMeanReference] bit NOT NULL
                                CONSTRAINT [DF_ChartDefinition_UseMeanReference]
                                DEFAULT (1),

                            [IsEnabled] bit NOT NULL
                                CONSTRAINT [DF_ChartDefinition_IsEnabled]
                                DEFAULT (1),
                            [ChartOrder] int NOT NULL,

                            [ReportPlacementReference] nvarchar(200) NULL,

                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_ChartDefinition_IsDeleted]
                                DEFAULT (0),

                            [CreatedUtc] datetime2(3) NOT NULL
                                CONSTRAINT [DF_ChartDefinition_CreatedUtc]
                                DEFAULT (SYSUTCDATETIME()),
                            [UpdatedUtc] datetime2(3) NOT NULL
                                CONSTRAINT [DF_ChartDefinition_UpdatedUtc]
                                DEFAULT (SYSUTCDATETIME()),

                            CONSTRAINT [PK_ChartDefinition]
                                PRIMARY KEY CLUSTERED ([ChartDefinition_ID]),

                            CONSTRAINT [UQ_ChartDefinition_Guid]
                                UNIQUE ([ChartDefinitionGuid]),

                            CONSTRAINT [UQ_ChartDefinition_Project_ChartNumber]
                                UNIQUE ([Project_ID], [ChartNumber]),

                            CONSTRAINT [FK_ChartDefinition_Project]
                                FOREIGN KEY ([Project_ID])
                                REFERENCES [dbo].[Project] ([Project_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION,

                            CONSTRAINT [FK_ChartDefinition_ChartType]
                                FOREIGN KEY ([ChartType_ID])
                                REFERENCES [dbo].[ChartType] ([ChartType_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION,

                            CONSTRAINT [CK_ChartDefinition_ChartNumber]
                                CHECK ([ChartNumber] > 0),

                            CONSTRAINT [CK_ChartDefinition_ChartOrder]
                                CHECK ([ChartOrder] > 0),

                            CONSTRAINT [CK_ChartDefinition_PngDimensions]
                                CHECK
                                (
                                    [PngWidthPixels] > 0
                                    AND [PngHeightPixels] > 0
                                ),

                            CONSTRAINT [CK_ChartDefinition_YAxis]
                                CHECK
                                (
                                    [UseAutomaticYAxis] = 1
                                    OR
                                    (
                                        [FixedYAxisMinimum] IS NOT NULL
                                        AND [FixedYAxisMaximum] IS NOT NULL
                                        AND [FixedYAxisMinimum] < [FixedYAxisMaximum]
                                    )
                                ),

                            CONSTRAINT [CK_ChartDefinition_TimeWindowMode]
                                CHECK
                                (
                                    [ChartTimeWindowMode] IN
                                    (
                                        N'Absolute',
                                        N'Relative'
                                    )
                                ),

                            CONSTRAINT [CK_ChartDefinition_RelativeAnchor]
                                CHECK
                                (
                                    [RelativeAnchor] IS NULL
                                    OR [RelativeAnchor] IN
                                    (
                                        N'ReportStart',
                                        N'ReportEnd'
                                    )
                                ),

                            CONSTRAINT [CK_ChartDefinition_ReferenceMode]
                                CHECK
                                (
                                    [ReferenceMode] IN
                                    (
                                        N'None',
                                        N'FixedDateTime'
                                    )
                                ),

                            CONSTRAINT [CK_ChartDefinition_TimeWindowValues]
                                CHECK
                                (
                                    (
                                        [ChartTimeWindowMode] = N'Absolute'
                                        AND [AbsoluteStartUtc] IS NOT NULL
                                        AND [AbsoluteEndUtc] IS NOT NULL
                                        AND [AbsoluteStartUtc] < [AbsoluteEndUtc]
                                    )
                                    OR
                                    (
                                        [ChartTimeWindowMode] = N'Relative'
                                        AND [RelativeAnchor] IS NOT NULL
                                        AND [RelativeStartOffsetSec] IS NOT NULL
                                        AND [RelativeEndOffsetSec] IS NOT NULL
                                        AND [RelativeStartOffsetSec] < [RelativeEndOffsetSec]
                                    )
                                ),

                            CONSTRAINT [CK_ChartDefinition_ReferenceValues]
                                CHECK
                                (
                                    [ReferenceMode] = N'None'
                                    OR
                                    (
                                        [ReferenceMode] = N'FixedDateTime'
                                        AND [ReferenceDateTimeUtc] IS NOT NULL
                                        AND [ReferenceBlockSeconds] IS NOT NULL
                                        AND [ReferenceBlockSeconds] > 0
                                    )
                                )
                        );

                        CREATE INDEX [IX_ChartDefinition_Project_Order]
                            ON [dbo].[ChartDefinition]
                            (
                                [Project_ID],
                                [IsDeleted],
                                [IsEnabled],
                                [ChartOrder]
                            );

                        CREATE UNIQUE INDEX
                            [UX_ChartDefinition_Active_Project_ChartName]
                            ON [dbo].[ChartDefinition]
                            (
                                [Project_ID],
                                [ChartName]
                            )
                            WHERE [IsDeleted] = 0;

                    END;

                    /* =============================================================
                       CHART DEFINITION SERIES
                       One selected monitored entity/data element plotted on a
                       committed chart.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ChartDefinitionSeries', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ChartDefinitionSeries]
                        (
                            [ChartDefinitionSeries_ID] int IDENTITY(1,1) NOT NULL,
                            [ChartDefinition_ID] int NOT NULL,
                            [EntityId] int NOT NULL,
                            [EntityDisplayName] nvarchar(200) NOT NULL,
                            [DataElementKey] nvarchar(100) NOT NULL,
                            [LegendText] nvarchar(300) NOT NULL,
                            [ColourHex] nvarchar(20) NULL,
                            [LineWidth] decimal(10,3) NOT NULL
                                CONSTRAINT [DF_ChartDefinitionSeries_LineWidth]
                                DEFAULT (2.0),
                            [MarkerSize] decimal(10,3) NOT NULL
                                CONSTRAINT [DF_ChartDefinitionSeries_MarkerSize]
                                DEFAULT (4.0),
                            [DisplayOrder] int NOT NULL,

                            CONSTRAINT [PK_ChartDefinitionSeries]
                                PRIMARY KEY CLUSTERED ([ChartDefinitionSeries_ID]),

                            CONSTRAINT [FK_ChartDefinitionSeries_ChartDefinition]
                                FOREIGN KEY ([ChartDefinition_ID])
                                REFERENCES [dbo].[ChartDefinition] ([ChartDefinition_ID])
                                ON DELETE CASCADE
                                ON UPDATE NO ACTION,

                            CONSTRAINT [UQ_ChartDefinitionSeries_Order]
                                UNIQUE ([ChartDefinition_ID], [DisplayOrder]),

                            CONSTRAINT [CK_ChartDefinitionSeries_EntityId]
                                CHECK ([EntityId] > 0),

                            CONSTRAINT [CK_ChartDefinitionSeries_DisplayOrder]
                                CHECK ([DisplayOrder] > 0),

                            CONSTRAINT [CK_ChartDefinitionSeries_LineWidth]
                                CHECK ([LineWidth] > 0),

                            CONSTRAINT [CK_ChartDefinitionSeries_MarkerSize]
                                CHECK ([MarkerSize] >= 0)
                        );

                        CREATE INDEX [IX_ChartDefinitionSeries_Definition]
                            ON [dbo].[ChartDefinitionSeries]
                            (
                                [ChartDefinition_ID],
                                [DisplayOrder]
                            );

                    END;

                    /* =============================================================
                       CHART DEFINITION TRIGGER BAND
                       Persistent Green / Yellow / Amber / Red chart background
                       configuration.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ChartDefinitionTriggerBand', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ChartDefinitionTriggerBand]
                        (
                            [ChartDefinitionTriggerBand_ID] int IDENTITY(1,1) NOT NULL,
                            [ChartDefinition_ID] int NOT NULL,
                            [BandName] nvarchar(100) NOT NULL,
                            [MinimumValue] decimal(18,6) NOT NULL,
                            [MaximumValue] decimal(18,6) NOT NULL,
                            [IsSymmetric] bit NOT NULL
                                CONSTRAINT [DF_ChartDefinitionTriggerBand_IsSymmetric]
                                DEFAULT (0),
                            [ColourHex] nvarchar(20) NOT NULL,
                            [Opacity] decimal(6,5) NOT NULL,
                            [DrawBoundaryLines] bit NOT NULL
                                CONSTRAINT [DF_ChartDefinitionTriggerBand_DrawBoundaryLines]
                                DEFAULT (0),
                            [DisplayOrder] int NOT NULL,

                            CONSTRAINT [PK_ChartDefinitionTriggerBand]
                                PRIMARY KEY CLUSTERED ([ChartDefinitionTriggerBand_ID]),

                            CONSTRAINT [FK_ChartDefinitionTriggerBand_ChartDefinition]
                                FOREIGN KEY ([ChartDefinition_ID])
                                REFERENCES [dbo].[ChartDefinition] ([ChartDefinition_ID])
                                ON DELETE CASCADE
                                ON UPDATE NO ACTION,

                            CONSTRAINT [UQ_ChartDefinitionTriggerBand_Order]
                                UNIQUE ([ChartDefinition_ID], [DisplayOrder]),

                            CONSTRAINT [CK_ChartDefinitionTriggerBand_Range]
                                CHECK ([MinimumValue] < [MaximumValue]),

                            CONSTRAINT [CK_ChartDefinitionTriggerBand_Opacity]
                                CHECK
                                (
                                    [Opacity] >= 0
                                    AND [Opacity] <= 1
                                ),

                            CONSTRAINT [CK_ChartDefinitionTriggerBand_DisplayOrder]
                                CHECK ([DisplayOrder] > 0)
                        );

                    END;

                    /* Additive chart-editor migration. Legacy date-window columns
                       remain temporarily for compatibility with existing databases. */
                    IF COL_LENGTH(N'dbo.ChartDefinition', N'StartDateMode') IS NULL
                    BEGIN
                        ALTER TABLE [dbo].[ChartDefinition]
                            ADD [StartDateMode] nvarchar(20) NOT NULL
                                CONSTRAINT [DF_ChartDefinition_StartDateMode]
                                DEFAULT (N'ReportStart');
                    END;

                    IF COL_LENGTH(N'dbo.ChartDefinition', N'WidthMm') IS NULL
                    BEGIN
                        ALTER TABLE [dbo].[ChartDefinition]
                            ADD [WidthMm] int NOT NULL
                                    CONSTRAINT [DF_ChartDefinition_WidthMm] DEFAULT (150),
                                [HeightMm] int NOT NULL
                                    CONSTRAINT [DF_ChartDefinition_HeightMm] DEFAULT (80),
                                [ResolutionDpi] smallint NOT NULL
                                    CONSTRAINT [DF_ChartDefinition_ResolutionDpi] DEFAULT (300),
                                [ChartFontFamily] nvarchar(100) NOT NULL
                                    CONSTRAINT [DF_ChartDefinition_ChartFontFamily] DEFAULT (N'Arial');
                    END;

                    /* =============================================================
                       CHART TEMPLATE
                       Reusable chart canvas configuration.
                       Templates do not consume ChartDefinition.ChartNumber.
                       Actual monitored entities and permanent report dates are
                       deliberately excluded from the template.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ChartTemplate', N'U') IS NULL
                    BEGIN
                        CREATE TABLE [dbo].[ChartTemplate]
                        (
                            [ChartTemplate_ID] int IDENTITY(1,1) NOT NULL,
                            [Project_ID] int NOT NULL,
                            [TemplateName] nvarchar(200) NOT NULL,
                            [ChartType_ID] int NOT NULL,
                            [DefaultDurationDays] int NOT NULL
                                CONSTRAINT [DF_ChartTemplate_DefaultDurationDays]
                                DEFAULT (14),
                            [GreenTrigger] decimal(18,6) NOT NULL,
                            [AmberTrigger] decimal(18,6) NOT NULL,
                            [RedTrigger] decimal(18,6) NOT NULL,
                            [UseAutomaticVerticalAxis] bit NOT NULL
                                CONSTRAINT [DF_ChartTemplate_UseAutomaticVerticalAxis]
                                DEFAULT (1),
                            [VerticalAxisMinimum] decimal(18,6) NOT NULL,
                            [VerticalAxisMaximum] decimal(18,6) NOT NULL,
                            [ChartTitle] nvarchar(500) NULL,
                            [VerticalAxisTitle] nvarchar(200) NOT NULL,
                            [YAxisUnit] nvarchar(50) NOT NULL,
                            [ShowLegend] bit NOT NULL
                                CONSTRAINT [DF_ChartTemplate_ShowLegend]
                                DEFAULT (1),
                            [LegendPosition] nvarchar(50) NOT NULL,
                            [PngWidthPixels] int NOT NULL,
                            [PngHeightPixels] int NOT NULL,
                            [DefaultLineWidth] decimal(10,3) NOT NULL,
                            [DefaultMarkerSize] decimal(10,3) NOT NULL,
                            [ShowGridLines] bit NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_ChartTemplate_IsDeleted]
                                DEFAULT (0),
                            [CreatedUtc] datetime2(3) NOT NULL
                                CONSTRAINT [DF_ChartTemplate_CreatedUtc]
                                DEFAULT (SYSUTCDATETIME()),
                            [UpdatedUtc] datetime2(3) NOT NULL
                                CONSTRAINT [DF_ChartTemplate_UpdatedUtc]
                                DEFAULT (SYSUTCDATETIME()),

                            CONSTRAINT [PK_ChartTemplate]
                                PRIMARY KEY CLUSTERED ([ChartTemplate_ID]),

                            CONSTRAINT [FK_ChartTemplate_Project]
                                FOREIGN KEY ([Project_ID])
                                REFERENCES [dbo].[Project] ([Project_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION,

                            CONSTRAINT [FK_ChartTemplate_ChartType]
                                FOREIGN KEY ([ChartType_ID])
                                REFERENCES [dbo].[ChartType] ([ChartType_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION,

                            CONSTRAINT [CK_ChartTemplate_Duration]
                                CHECK ([DefaultDurationDays] > 0),

                            CONSTRAINT [CK_ChartTemplate_Triggers]
                                CHECK
                                (
                                    [GreenTrigger] > 0
                                    AND [GreenTrigger] < [AmberTrigger]
                                    AND [AmberTrigger] < [RedTrigger]
                                ),

                            CONSTRAINT [CK_ChartTemplate_VerticalAxis]
                                CHECK ([VerticalAxisMinimum] < [VerticalAxisMaximum])
                        );

                        CREATE UNIQUE INDEX
                            [UX_ChartTemplate_Active_Project_Name]
                            ON [dbo].[ChartTemplate]
                            (
                                [Project_ID],
                                [TemplateName]
                            )
                            WHERE [IsDeleted] = 0;
                    END;

                    IF COL_LENGTH(N'dbo.ChartTemplate', N'DefaultStartDateMode') IS NULL
                    BEGIN
                        ALTER TABLE [dbo].[ChartTemplate]
                            ADD [DefaultStartDateMode] nvarchar(20) NOT NULL
                                    CONSTRAINT [DF_ChartTemplate_DefaultStartDateMode]
                                    DEFAULT (N'ReportStart'),
                                [WidthMm] int NOT NULL
                                    CONSTRAINT [DF_ChartTemplate_WidthMm] DEFAULT (150),
                                [HeightMm] int NOT NULL
                                    CONSTRAINT [DF_ChartTemplate_HeightMm] DEFAULT (80),
                                [ResolutionDpi] smallint NOT NULL
                                    CONSTRAINT [DF_ChartTemplate_ResolutionDpi] DEFAULT (300),
                                [ChartFontFamily] nvarchar(100) NOT NULL
                                    CONSTRAINT [DF_ChartTemplate_ChartFontFamily] DEFAULT (N'Arial');
                    END;


                    /* =============================================================
                       CHART TEMPLATE SERIES
                       Series definitions stored without monitored entity identity.
                       Entity selection is supplied only when a real chart is made.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ChartTemplateSeries', N'U') IS NULL
                    BEGIN
                        CREATE TABLE [dbo].[ChartTemplateSeries]
                        (
                            [ChartTemplateSeries_ID] int IDENTITY(1,1) NOT NULL,
                            [ChartTemplate_ID] int NOT NULL,
                            [DataElementKey] nvarchar(100) NOT NULL,
                            [LegendText] nvarchar(300) NULL,
                            [LineWidth] decimal(10,3) NOT NULL,
                            [MarkerSize] decimal(10,3) NOT NULL,
                            [DisplayOrder] int NOT NULL,

                            CONSTRAINT [PK_ChartTemplateSeries]
                                PRIMARY KEY CLUSTERED ([ChartTemplateSeries_ID]),

                            CONSTRAINT [FK_ChartTemplateSeries_ChartTemplate]
                                FOREIGN KEY ([ChartTemplate_ID])
                                REFERENCES [dbo].[ChartTemplate] ([ChartTemplate_ID])
                                ON DELETE CASCADE
                                ON UPDATE NO ACTION,

                            CONSTRAINT [UQ_ChartTemplateSeries_Order]
                                UNIQUE ([ChartTemplate_ID], [DisplayOrder]),

                            CONSTRAINT [CK_ChartTemplateSeries_DisplayOrder]
                                CHECK ([DisplayOrder] > 0)
                        );
                    END;


                    COMMIT TRANSACTION;

                END TRY

                BEGIN CATCH

                    IF @@TRANCOUNT > 0
                        ROLLBACK TRANSACTION;

                    THROW;

                END CATCH;
                """;

            #endregion


            #region Create Missing Tables

            await using (SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionString))
            {
                await databaseConnection.OpenAsync();

                await using SqlCommand createTablesCommand =
                    new(
                        cmdText: createTablesSql,
                        connection: databaseConnection);

                await createTablesCommand.ExecuteNonQueryAsync();
            }

            #endregion



        }

        #endregion


        #region Project Configuration

        #region Cross-Instance Project Lock Infrastructure

        private static string BuildProjectLockResourceName(
            int projectId)
        {
            #region Validate Project ID

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(projectId),
                    message: "Project_ID must be greater than zero.");
            }

            #endregion


            #region Build Lock Resource Name

            return
                $"{ProjectLockResourcePrefix}{projectId}";

            #endregion
        }


        private async Task<SqlConnection> OpenSharedProjectLockConnectionAsync(
            int projectId)
        {
            #region Validate Project ID

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(projectId),
                    message: "Project_ID must be greater than zero.");
            }

            #endregion


            #region Build Dedicated Lock Connection String

            string lockConnectionString =
                BuildDatabaseConnectionString(
                    baseConnectionString: GetValidatedDatabaseConnectionString(),
                    databaseName: TrackGeometryDatabaseName,
                    pooling: false);

            #endregion


            #region Open Dedicated Lock Connection

            SqlConnection lockConnection =
                new(
                    connectionString:
                        lockConnectionString);

            try
            {
                await lockConnection.OpenAsync();

            #endregion


                #region Acquire Shared Project Lock

                const string acquireSharedLockSql = """
            DECLARE @LockResult int;

            EXEC @LockResult = sys.sp_getapplock
                @Resource = @Resource,
                @LockMode = 'Shared',
                @LockOwner = 'Session',
                @LockTimeout = 0;

            SELECT @LockResult;
            """;

                await using SqlCommand acquireSharedLockCommand =
                    new(
                        cmdText: acquireSharedLockSql,
                        connection: lockConnection);

                acquireSharedLockCommand.Parameters.AddWithValue(
                    parameterName: "@Resource",
                    value: BuildProjectLockResourceName(
                        projectId: projectId));

                object lockResultObject =
                    await acquireSharedLockCommand.ExecuteScalarAsync()
                    ?? throw new InvalidOperationException(
                        "SQL Server returned no result when acquiring the project lock.");

                int lockResult =
                    Convert.ToInt32(
                        value: lockResultObject);

                if (lockResult < 0)
                {
                    throw new InvalidOperationException(
                        $"Project_ID {projectId} is temporarily unavailable because " +
                        "another application instance is changing its project state.");
                }

                #endregion


                #region Return Locked Connection

                // This connection must remain open for as long as this project is
                // active in this application instance.

                return lockConnection;

                #endregion
            }
            catch
            {
                #region Release Failed Lock Connection

                await lockConnection.DisposeAsync();

                #endregion

                throw;
            }
        }


        #region Exclusive Project Lock

        private async Task<SqlConnection> OpenExclusiveProjectLockConnectionAsync(
            int projectId)
        {
            #region Validate Project ID

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(projectId),
                    message: "Project_ID must be greater than zero.");
            }

            #endregion


            #region Build Dedicated Lock Connection String

            string lockConnectionString =
                BuildDatabaseConnectionString(
                    baseConnectionString: GetValidatedDatabaseConnectionString(),
                    databaseName: TrackGeometryDatabaseName,
                    pooling: false);

            #endregion


            #region Open Dedicated Lock Connection

            SqlConnection lockConnection =
                new(
                    connectionString:
                        lockConnectionString);

            try
            {
                await lockConnection.OpenAsync();

            #endregion


                #region Acquire Exclusive Project Lock

                const string acquireExclusiveLockSql = """
            DECLARE @LockResult int;

            EXEC @LockResult = sys.sp_getapplock
                @Resource = @Resource,
                @LockMode = 'Exclusive',
                @LockOwner = 'Session',
                @LockTimeout = 0;

            SELECT @LockResult;
            """;

                await using SqlCommand acquireExclusiveLockCommand =
                    new(
                        cmdText: acquireExclusiveLockSql,
                        connection: lockConnection);

                acquireExclusiveLockCommand.Parameters.AddWithValue(
                    parameterName: "@Resource",
                    value: BuildProjectLockResourceName(
                        projectId: projectId));

                object lockResultObject =
                    await acquireExclusiveLockCommand.ExecuteScalarAsync()
                    ?? throw new InvalidOperationException(
                        "SQL Server returned no result when acquiring the " +
                        "exclusive project lock.");

                int lockResult =
                    Convert.ToInt32(
                        value: lockResultObject);

                if (lockResult < 0)
                {
                    throw new InvalidOperationException(
                        $"Project_ID {projectId} cannot be deleted because it is " +
                        "currently active or otherwise in use by another running " +
                        "application instance.");
                }

                #endregion


                #region Return Exclusively Locked Connection

                // The caller must keep this connection open throughout the complete
                // delete operation.
                //
                // Disposing the connection releases the Exclusive application lock.

                return lockConnection;

                #endregion
            }
            catch
            {
                #region Dispose Failed Lock Connection

                await lockConnection.DisposeAsync();

                #endregion

                throw;
            }
        }

        #endregion




        #region Initialise Startup Active Project

        private async Task InitialiseStartupActiveProjectAsync()
        {
            #region Handle No Persisted Active Project

            if (!_activeProjectId.HasValue)
            {
                _activeProjectName =
                    string.Empty;

                txtActiveProject.Text =
                    "No active project";

                return;
            }

            #endregion


            #region Capture Startup Project ID

            // Capture the value once.
            // Do not re-read the Registry during this operation.

            int startupProjectId =
                _activeProjectId.Value;

            SqlConnection? startupLockConnection =
                null;

            #endregion


            try
            {
                #region Acquire Shared Project Lock

                startupLockConnection =
                    await OpenSharedProjectLockConnectionAsync(
                        projectId: startupProjectId);

                #endregion


                #region Define Project Validation Query

                const string validateProjectSql = """
            SELECT
                [ProjectName],
                [IsDeleted]
            FROM [dbo].[Project]
            WHERE [Project_ID] = @Project_ID;
            """;

                #endregion


                #region Validate Project Against Database

                await using SqlCommand validateProjectCommand =
                    new(
                        cmdText: validateProjectSql,
                        connection: startupLockConnection);

                validateProjectCommand.Parameters.AddWithValue(
                    parameterName: "@Project_ID",
                    value: startupProjectId);

                await using SqlDataReader reader =
                    await validateProjectCommand.ExecuteReaderAsync();

                if (!await reader.ReadAsync())
                {
                    #region Handle Missing Project

                    await startupLockConnection.DisposeAsync();

                    startupLockConnection =
                        null;

                    _activeProjectId =
                        null;

                    _activeProjectName =
                        string.Empty;

                    txtActiveProject.Text =
                        "No active project";

                    return;

                    #endregion
                }


                string databaseProjectName =
                    reader.GetString(
                        i: 0);

                bool projectIsDeleted =
                    reader.GetBoolean(
                        i: 1);

                #endregion


                #region Reject Deleted Project

                if (projectIsDeleted)
                {
                    await startupLockConnection.DisposeAsync();

                    startupLockConnection =
                        null;

                    _activeProjectId =
                        null;

                    _activeProjectName =
                        string.Empty;

                    txtActiveProject.Text =
                        "No active project";

                    return;
                }

                #endregion


                #region Establish Process-Local Active Project

                // Ownership of this open connection now transfers to the running
                // application instance.
                //
                // The connection must remain open for as long as this project
                // remains active in this process.

                _activeProjectLockConnection =
                    startupLockConnection;

                startupLockConnection =
                    null;

                #endregion


                #region Refresh Runtime Project Name

                // Project_ID is authoritative.
                //
                // The SQL ProjectName replaces the Registry copy in this process
                // in case the project has subsequently been renamed.
                //
                // Do NOT write the refreshed name back to the Registry here.
                // Another running instance may have changed the startup default.

                _activeProjectName =
                    databaseProjectName;

                txtActiveProject.Text =
                    _activeProjectName;

                #endregion
            }
            catch (Exception ex)
            {
                #region Dispose Incomplete Lock Connection

                if (startupLockConnection is not null)
                {
                    await startupLockConnection.DisposeAsync();

                    startupLockConnection =
                        null;
                }

                #endregion


                #region Clear Process-Local Active Project

                ReleaseActiveProjectLockConnection();

                _activeProjectId =
                    null;

                _activeProjectName =
                    string.Empty;

                txtActiveProject.Text =
                    "No active project";

                #endregion


                #region Report Startup Lock Failure

                // Do not clear or modify the Registry here.
                //
                // The failure may be temporary and another running instance may
                // have changed the persisted startup default since this process
                // originally read it.

                txtDbConnectionStatus.Text =
                    $"Startup active project unavailable: {ex.Message}";

                #endregion
            }
        }

        #endregion




        private void ReleaseActiveProjectLockConnection()
        {
            #region Release Current Active Project Lock

            if (_activeProjectLockConnection is not null)
            {
                // Pooling is disabled for this dedicated connection.
                // Disposing it closes the underlying SQL session and therefore
                // releases the Session-owned application lock.

                _activeProjectLockConnection.Dispose();

                _activeProjectLockConnection =
                    null;
            }

            #endregion
        }

        #endregion






        private async void btnManageProjects_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Display Project Management Overlay

            brdProjectManagement.Visibility =
                Visibility.Visible;

            txtProjectManagementStatus.Text =
                "Loading projects...";

            #endregion


            #region Load Project Data

            try
            {
                await LoadProjectsAsync();

                #region Populate Active Project Edit Controls

                if (_activeProjectId.HasValue)
                {
                    ProjectConfigurationItem? activeProject =
                        null;

                    foreach (ProjectConfigurationItem projectItem
                        in _projectItems)
                    {
                        if (projectItem.Project_ID == _activeProjectId.Value)
                        {
                            activeProject =
                                projectItem;

                            break;
                        }
                    }

                    if (activeProject is not null)
                    {
                        dgProjects.SelectedItem =
                            activeProject;

                        txtNewProjectName.Text =
                            activeProject.ProjectName;

                        if (_projectStartDates.TryGetValue(
                            key: activeProject.Project_ID,
                            value: out DateTime activeProjectStartDate))
                        {
                            _pendingProjectStartDate =
                                activeProjectStartDate.Date;

                            _isUpdatingProjectStartDateControl =
                                true;

                            try
                            {
                                dpProjectStartDate.SelectedDate =
                                    _pendingProjectStartDate;
                            }
                            finally
                            {
                                _isUpdatingProjectStartDateControl =
                                    false;
                            }
                        }
                    }
                }

                #endregion

                txtProjectManagementStatus.Text =
                    $"{_projectItems.Count} project(s) loaded.";
            }
            catch (SqlException ex)
            {
                txtProjectManagementStatus.Text =
                    $"Unable to load projects: {ex.Message}";
            }
            catch (ArgumentException ex)
            {
                txtProjectManagementStatus.Text =
                    $"Invalid database configuration: {ex.Message}";
            }
            catch (InvalidOperationException ex)
            {
                txtProjectManagementStatus.Text =
                    $"Unable to load projects: {ex.Message}";
            }
            catch (Exception ex)
            {
                txtProjectManagementStatus.Text =
                    $"Unable to load projects: {ex.Message}";
            }

            #endregion
        }


        private void btnCloseProjectManagement_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Hide Project Management Overlay

            brdProjectManagement.Visibility =
                Visibility.Collapsed;

            #endregion
        }




        #region Project Start Date

        private async void dgProjects_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Synchronise Project Start Date Picker

            if (_isUpdatingProjectStartDateControl)
            {
                return;
            }

            if (dgProjects.SelectedItem
                is not ProjectConfigurationItem selectedProject)
            {
                return;
            }

            if (!_projectStartDates.TryGetValue(
                key: selectedProject.Project_ID,
                value: out DateTime projectStartDate))
            {
                projectStartDate =
                    DateTime.Today;
            }

            _pendingProjectStartDate =
                projectStartDate.Date;

            _isUpdatingProjectStartDateControl =
                true;

            try
            {
                dpProjectStartDate.SelectedDate =
                    _pendingProjectStartDate;
            }
            finally
            {
                _isUpdatingProjectStartDateControl =
                    false;
            }

            #endregion

            try
            {
                await LoadSelectedProjectSettingsAsync(
                    projectId: selectedProject.Project_ID);
            }
            catch (Exception ex)
            {
                txtProjectManagementStatus.Text =
                    $"Unable to load project settings: {ex.Message}";
            }
        }


        private async Task LoadSelectedProjectSettingsAsync(
            int projectId)
        {
            const string sql = """
                SELECT [TimeZoneId], [DefaultReportOutputPath]
                FROM [dbo].[Project]
                WHERE [Project_ID] = @Project_ID;
                """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString: GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand command =
                new(
                    cmdText: sql,
                    connection: databaseConnection);

            command.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value = projectId;

            await using SqlDataReader reader =
                await command.ExecuteReaderAsync();

            if (!await reader.ReadAsync())
            {
                throw new InvalidOperationException(
                    "The selected project no longer exists.");
            }

            string timeZoneId =
                reader.IsDBNull(0)
                    ? string.Empty
                    : reader.GetString(0);

            txtProjectOutputPath.Text =
                reader.IsDBNull(1)
                    ? string.Empty
                    : reader.GetString(1);

            cmbProjectTimeZone.SelectedValue =
                timeZoneId;

            txtChartProjectTimeZone.Text =
                string.IsNullOrWhiteSpace(timeZoneId)
                    ? "Not configured"
                    : timeZoneId;
        }


        private async void btnSaveProjectSettings_Click(
            object sender,
            RoutedEventArgs e)
        {
            if (dgProjects.SelectedItem
                is not ProjectConfigurationItem selectedProject)
            {
                txtProjectManagementStatus.Text =
                    "Select a project before saving project settings.";

                return;
            }

            string? timeZoneId =
                cmbProjectTimeZone.SelectedValue?.ToString();

            if (string.IsNullOrWhiteSpace(timeZoneId))
            {
                txtProjectManagementStatus.Text =
                    "Select a project time zone.";

                return;
            }

            string outputPath =
                txtProjectOutputPath.Text.Trim();

            const string sql = """
                UPDATE [dbo].[Project]
                SET
                    [TimeZoneId] = @TimeZoneId,
                    [DefaultReportOutputPath] = NULLIF(@DefaultReportOutputPath, N'')
                WHERE [Project_ID] = @Project_ID;
                """;

            try
            {
                _ = TimeZoneInfo.FindSystemTimeZoneById(
                    id: timeZoneId);

                await using SqlConnection databaseConnection =
                    new(
                        connectionString: GetTrackGeometryConnectionString());

                await databaseConnection.OpenAsync();

                await using SqlCommand command =
                    new(
                        cmdText: sql,
                        connection: databaseConnection);

                command.Parameters.Add(
                    parameterName: "@Project_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value = selectedProject.Project_ID;

                command.Parameters.Add(
                    parameterName: "@TimeZoneId",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 200)
                    .Value = timeZoneId;

                command.Parameters.Add(
                    parameterName: "@DefaultReportOutputPath",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 1000)
                    .Value = outputPath;

                int affectedRows =
                    await command.ExecuteNonQueryAsync();

                if (affectedRows != 1)
                {
                    throw new InvalidOperationException(
                        "The selected project changed before its settings could be saved.");
                }

                txtChartProjectTimeZone.Text =
                    timeZoneId;

                txtProjectManagementStatus.Text =
                    "Project time zone and default report output path saved.";
            }
            catch (Exception ex)
            {
                txtProjectManagementStatus.Text =
                    $"Unable to save project settings: {ex.Message}";
            }
        }


        private async void dpProjectStartDate_SelectedDateChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            // This handler is intentionally available for future XAML use.
            // The current DatePicker is used both for a new project and for an
            // existing selected project. Existing-project persistence occurs
            // explicitly through the calendar-selection workflow below.
            await Task.CompletedTask;
        }


        private void ProjectStartDateTextBlock_Loaded(
            object sender,
            RoutedEventArgs e)
        {
            #region Populate Project Start Date Cell

            if (sender is not TextBlock dateTextBlock ||
                dateTextBlock.DataContext
                    is not ProjectConfigurationItem projectItem)
            {
                return;
            }

            if (_projectStartDates.TryGetValue(
                key: projectItem.Project_ID,
                value: out DateTime projectStartDate))
            {
                dateTextBlock.Text =
                    projectStartDate.ToString(
                        format: "yyyy-MM-dd",
                        provider: CultureInfo.InvariantCulture);
            }
            else
            {
                dateTextBlock.Text =
                    DateTime.Today.ToString(
                        format: "yyyy-MM-dd",
                        provider: CultureInfo.InvariantCulture);
            }

            #endregion
        }


        private async Task UpdateSelectedProjectStartDateAsync(
            DateTime projectStartDate)
        {
            #region Validate Selected Project

            if (dgProjects.SelectedItem
                is not ProjectConfigurationItem selectedProject)
            {
                _pendingProjectStartDate =
                    projectStartDate.Date;

                txtProjectManagementStatus.Text =
                    $"Project start date for new project: {_pendingProjectStartDate:yyyy-MM-dd}.";

                return;
            }

            if (selectedProject.Project_ID <= 0)
            {
                throw new InvalidOperationException(
                    "Selected project is invalid.");
            }

            #endregion


            #region Update Database

            const string updateSql = """
                UPDATE [dbo].[Project]
                SET [ProjectStartDate] = @ProjectStartDate
                WHERE [Project_ID] = @Project_ID;
                """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand updateCommand =
                new(
                    cmdText: updateSql,
                    connection: databaseConnection);

            updateCommand.Parameters.Add(
                parameterName: "@ProjectStartDate",
                sqlDbType: System.Data.SqlDbType.Date)
                .Value =
                    projectStartDate.Date;

            updateCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    selectedProject.Project_ID;

            int affectedRows =
                await updateCommand.ExecuteNonQueryAsync();

            if (affectedRows != 1)
            {
                throw new InvalidOperationException(
                    "Project start date was not updated.");
            }

            #endregion


            #region Refresh Runtime State

            _projectStartDates[selectedProject.Project_ID] =
                projectStartDate.Date;

            _pendingProjectStartDate =
                projectStartDate.Date;

            dgProjects.Items.Refresh();

            txtProjectManagementStatus.Text =
                $"Project '{selectedProject.ProjectName}' start date set to " +
                $"{projectStartDate:yyyy-MM-dd}.";

            #endregion
        }


        private async void btnProjectStartCalendar_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Resolve Project Start Calendar Target

            bool creatingNewProject =
                !string.IsNullOrWhiteSpace(
                    value: txtNewProjectName.Text);

            if (!creatingNewProject &&
                dgProjects.SelectedItem
                    is ProjectConfigurationItem selectedProject)
            {
                _projectStartCalendarTargetProjectId =
                    selectedProject.Project_ID;

                if (_projectStartDates.TryGetValue(
                    key: selectedProject.Project_ID,
                    value: out DateTime existingStartDate))
                {
                    dpProjectStartDate.SelectedDate =
                        existingStartDate.Date;
                }
                else
                {
                    dpProjectStartDate.SelectedDate =
                        DateTime.Today;
                }
            }
            else
            {
                _projectStartCalendarTargetProjectId =
                    null;

                dpProjectStartDate.SelectedDate =
                    _pendingProjectStartDate.Date;
            }

            dpProjectStartDate.Focus();

            dpProjectStartDate.IsDropDownOpen =
                true;

            #endregion
        }


        private async void dpProjectStartDate_CalendarClosed(
            object? sender,
            RoutedEventArgs e)
        {
            #region Read Selected Project Start Date

            DateTime selectedDate =
                dpProjectStartDate.SelectedDate?.Date
                ?? DateTime.Today;

            #endregion


            #region Apply To Existing Project Or New Project

            try
            {
                if (_projectStartCalendarTargetProjectId.HasValue)
                {
                    int targetProjectId =
                        _projectStartCalendarTargetProjectId.Value;

                    ProjectConfigurationItem? targetProject =
                        null;

                    foreach (ProjectConfigurationItem projectItem
                        in _projectItems)
                    {
                        if (projectItem.Project_ID == targetProjectId)
                        {
                            targetProject =
                                projectItem;

                            break;
                        }
                    }

                    if (targetProject is null)
                    {
                        throw new InvalidOperationException(
                            "The selected project is no longer available.");
                    }

                    dgProjects.SelectedItem =
                        targetProject;

                    await UpdateSelectedProjectStartDateAsync(
                        projectStartDate: selectedDate);
                }
                else
                {
                    _pendingProjectStartDate =
                        selectedDate;

                    txtProjectManagementStatus.Text =
                        $"Project start date for new project: {selectedDate:yyyy-MM-dd}.";
                }
            }
            catch (Exception ex)
            {
                txtProjectManagementStatus.Text =
                    $"Unable to update project start date: {ex.Message}";
            }
            finally
            {
                _projectStartCalendarTargetProjectId =
                    null;
            }

            #endregion
        }

        #endregion


        #region Add Project

        private void ProjectRequiredField_Changed(
            object sender,
            RoutedEventArgs e)
        {
            UpdateAddProjectButtonState();
        }


        private void UpdateAddProjectButtonState()
        {
            if (!IsInitialized)
            {
                return;
            }

            ProjectConfigurationItem? matchingProject =
                FindProjectByEnteredName();

            btnAddProject.Content =
                matchingProject is null
                    ? "Add Project"
                    : "Overwrite";

            btnAddProject.IsEnabled =
                !string.IsNullOrWhiteSpace(txtNewProjectName.Text) &&
                dpProjectStartDate.SelectedDate.HasValue &&
                cmbProjectTimeZone.SelectedValue is string timeZoneId &&
                !string.IsNullOrWhiteSpace(timeZoneId) &&
                !string.IsNullOrWhiteSpace(txtProjectOutputPath.Text);
        }


        private ProjectConfigurationItem? FindProjectByEnteredName()
        {
            #region Resolve Entered Project Name

            string enteredProjectName =
                txtNewProjectName.Text?.Trim()
                ?? string.Empty;

            if (string.IsNullOrWhiteSpace(enteredProjectName))
            {
                return null;
            }

            #endregion


            #region Find Existing Project

            foreach (ProjectConfigurationItem projectItem
                in _projectItems)
            {
                if (string.Equals(
                    a: projectItem.ProjectName,
                    b: enteredProjectName,
                    comparisonType:
                        StringComparison.OrdinalIgnoreCase))
                {
                    return projectItem;
                }
            }

            return null;

            #endregion
        }


        private void btnSelectProjectOutputPath_Click(
            object sender,
            RoutedEventArgs e)
        {
            OpenFolderDialog dialog =
                new()
                {
                    Title = "Select Default Report Output Folder",
                    Multiselect = false
                };

            if (!string.IsNullOrWhiteSpace(txtProjectOutputPath.Text) &&
                Directory.Exists(
                    path: txtProjectOutputPath.Text))
            {
                dialog.InitialDirectory =
                    txtProjectOutputPath.Text;
            }

            if (dialog.ShowDialog(
                    owner: this) == true)
            {
                txtProjectOutputPath.Text =
                    dialog.FolderName;

                UpdateAddProjectButtonState();
            }
        }

        private async void btnAddProject_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Project Name

            string projectName =
                txtNewProjectName.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "New project name control returned null.");

            if (string.IsNullOrWhiteSpace(projectName))
            {
                txtProjectManagementStatus.Text =
                    "Enter a project name.";

                txtNewProjectName.Focus();

                return;
            }

            if (projectName.Length > 200)
            {
                txtProjectManagementStatus.Text =
                    "Project name cannot exceed 200 characters.";

                txtNewProjectName.Focus();

                return;
            }

            #endregion


            #region Prepare User Interface

            ProjectConfigurationItem? matchingProject =
                FindProjectByEnteredName();

            if (matchingProject is not null &&
                !matchingProject.IsDeleted)
            {
                MessageBoxResult overwriteConfirmation =
                    MessageBox.Show(
                        owner: this,
                        messageBoxText:
                            $"Overwrite the configuration values for project " +
                            $"'{matchingProject.ProjectName}'?\n\n" +
                            "The existing Project_ID and all associated project data " +
                            "will be retained.",
                        caption: "Confirm Project Overwrite",
                        button: MessageBoxButton.YesNo,
                        icon: MessageBoxImage.Question,
                        defaultResult: MessageBoxResult.No);

                if (overwriteConfirmation != MessageBoxResult.Yes)
                {
                    txtProjectManagementStatus.Text =
                        "Project overwrite cancelled.";

                    return;
                }
            }

            btnAddProject.IsEnabled =
                false;

            txtProjectManagementStatus.Text =
                $"Checking project '{projectName}'...";

            #endregion


            try
            {
                #region Add Or Restore Project

                DateTime projectStartDate =
                    dpProjectStartDate.SelectedDate?.Date
                    ?? throw new InvalidOperationException(
                        "Select a project start date.");

                string timeZoneId =
                    cmbProjectTimeZone.SelectedValue?.ToString()
                    ?? throw new InvalidOperationException(
                        "Select a project time zone.");

                string outputPath =
                    txtProjectOutputPath.Text.Trim();

                bool projectAddedOrRestored =
                    await AddOrRestoreProjectAsync(
                        projectName: projectName,
                        projectStartDate: projectStartDate,
                        timeZoneId: timeZoneId,
                        defaultReportOutputPath: outputPath);

                #endregion


                #region Refresh Project List

                if (projectAddedOrRestored)
                {
                    txtNewProjectName.Clear();

                    txtProjectOutputPath.Clear();

                    cmbProjectTimeZone.SelectedValue =
                        TimeZoneInfo.Utc.Id;

                    _pendingProjectStartDate =
                        DateTime.Today;

                    _isUpdatingProjectStartDateControl =
                        true;

                    try
                    {
                        dpProjectStartDate.SelectedDate =
                            _pendingProjectStartDate;
                    }
                    finally
                    {
                        _isUpdatingProjectStartDateControl =
                            false;
                    }

                    await LoadProjectsAsync();
                }

                #endregion
            }

            #region Handle Project Errors

            catch (SqlException ex)
                when (ex.Number == 2601 ||
                      ex.Number == 2627)
            {
                // SQL Server remains the final authority for uniqueness.
                // This also protects against two application instances attempting
                // to create the same project at approximately the same time.

                txtProjectManagementStatus.Text =
                    $"Project '{projectName}' already exists.";

                await LoadProjectsAsync();
            }
            catch (SqlException ex)
            {
                txtProjectManagementStatus.Text =
                    $"Unable to add project: {ex.Message}";
            }
            catch (ArgumentException ex)
            {
                txtProjectManagementStatus.Text =
                    $"Invalid project: {ex.Message}";
            }
            catch (InvalidOperationException ex)
            {
                txtProjectManagementStatus.Text =
                    $"Unable to add project: {ex.Message}";
            }
            catch (Exception ex)
            {
                txtProjectManagementStatus.Text =
                    $"Unable to add project: {ex.Message}";
            }

            #endregion


            #region Restore User Interface

            finally
            {
                UpdateAddProjectButtonState();
            }

            #endregion
        }


        private async Task<bool> AddOrRestoreProjectAsync(
            string projectName,
            DateTime projectStartDate,
            string timeZoneId,
            string defaultReportOutputPath)
        {
            #region Validate Project Name

            string validatedProjectName =
                projectName?.Trim()
                ?? throw new ArgumentNullException(
                    paramName: nameof(projectName));

            if (string.IsNullOrWhiteSpace(validatedProjectName))
            {
                throw new ArgumentException(
                    message: "Project name cannot be empty.",
                    paramName: nameof(projectName));
            }

            if (validatedProjectName.Length > 200)
            {
                throw new ArgumentException(
                    message: "Project name cannot exceed 200 characters.",
                    paramName: nameof(projectName));
            }

            DateTime validatedProjectStartDate =
                projectStartDate.Date;

            string validatedTimeZoneId =
                timeZoneId?.Trim()
                ?? throw new ArgumentNullException(
                    paramName: nameof(timeZoneId));

            _ = TimeZoneInfo.FindSystemTimeZoneById(
                id: validatedTimeZoneId);

            string validatedOutputPath =
                defaultReportOutputPath?.Trim()
                ?? throw new ArgumentNullException(
                    paramName: nameof(defaultReportOutputPath));

            if (string.IsNullOrWhiteSpace(validatedOutputPath))
            {
                throw new ArgumentException(
                    message: "Default report output path cannot be empty.",
                    paramName: nameof(defaultReportOutputPath));
            }

            #endregion


            #region Resolve Database Connection String

            string databaseConnectionString =
                GetTrackGeometryConnectionString();

            #endregion


            #region Open Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionString);

            await databaseConnection.OpenAsync();

            #endregion


            #region Check For Existing Project

            const string findProjectSql = """
        SELECT
            [Project_ID],
            [ProjectName],
            [IsDeleted]
        FROM [dbo].[Project]
        WHERE [ProjectName] = @ProjectName;
        """;

            await using SqlCommand findProjectCommand =
                new(
                    cmdText: findProjectSql,
                    connection: databaseConnection);

            findProjectCommand.Parameters.AddWithValue(
                parameterName: "@ProjectName",
                value: validatedProjectName);

            int? existingProjectId =
                null;

            bool existingProjectIsDeleted =
                false;

            await using (SqlDataReader reader =
                await findProjectCommand.ExecuteReaderAsync())
            {
                if (await reader.ReadAsync())
                {
                    existingProjectId =
                        reader.GetInt32(
                            i: 0);

                    existingProjectIsDeleted =
                        reader.GetBoolean(
                            i: 2);
                }
            }

            #endregion


            #region Handle Existing Active Project

            if (existingProjectId.HasValue &&
                !existingProjectIsDeleted)
            {
                const string updateProjectSql = """
                    UPDATE [dbo].[Project]
                    SET
                        [ProjectStartDate] = @ProjectStartDate,
                        [TimeZoneId] = @TimeZoneId,
                        [DefaultReportOutputPath] = @DefaultReportOutputPath
                    WHERE
                        [Project_ID] = @Project_ID
                        AND [IsDeleted] = 0;
                    """;

                await using SqlCommand updateProjectCommand =
                    new(
                        cmdText: updateProjectSql,
                        connection: databaseConnection);

                updateProjectCommand.Parameters.Add(
                    parameterName: "@Project_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value = existingProjectId.Value;

                updateProjectCommand.Parameters.Add(
                    parameterName: "@ProjectStartDate",
                    sqlDbType: System.Data.SqlDbType.Date)
                    .Value = validatedProjectStartDate;

                updateProjectCommand.Parameters.Add(
                    parameterName: "@TimeZoneId",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 200)
                    .Value = validatedTimeZoneId;

                updateProjectCommand.Parameters.Add(
                    parameterName: "@DefaultReportOutputPath",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 1000)
                    .Value = validatedOutputPath;

                int updatedRows =
                    await updateProjectCommand.ExecuteNonQueryAsync();

                if (updatedRows != 1)
                {
                    throw new InvalidOperationException(
                        $"Project '{validatedProjectName}' could not be overwritten " +
                        "because its database state changed before the operation completed.");
                }

                txtProjectManagementStatus.Text =
                    $"Project '{validatedProjectName}' overwritten successfully.";

                return true;
            }

            #endregion


            #region Handle Existing Deleted Project

            if (existingProjectId.HasValue &&
                existingProjectIsDeleted)
            {
                MessageBoxResult restoreConfirmation =
                    MessageBox.Show(
                        messageBoxText:
                            $"Project '{validatedProjectName}' already exists but is marked as deleted.\n\n" +
                            "Do you want to restore this project?",
                        caption: "Restore Project",
                        button: MessageBoxButton.YesNo,
                        icon: MessageBoxImage.Question,
                        defaultResult: MessageBoxResult.No);

                if (restoreConfirmation != MessageBoxResult.Yes)
                {
                    txtProjectManagementStatus.Text =
                        $"Project '{validatedProjectName}' was not restored.";

                    return false;
                }

                #region Restore Deleted Project

                const string restoreProjectSql = """
            UPDATE [dbo].[Project]
            SET
                [IsDeleted] = 0,
                [ProjectStartDate] = @ProjectStartDate,
                [TimeZoneId] = @TimeZoneId,
                [DefaultReportOutputPath] = @DefaultReportOutputPath
            WHERE [Project_ID] = @Project_ID
              AND [IsDeleted] = 1;
            """;

                await using SqlCommand restoreProjectCommand =
                    new(
                        cmdText: restoreProjectSql,
                        connection: databaseConnection);

                restoreProjectCommand.Parameters.AddWithValue(
                    parameterName: "@Project_ID",
                    value: existingProjectId.Value);

                restoreProjectCommand.Parameters.Add(
                    parameterName: "@ProjectStartDate",
                    sqlDbType: System.Data.SqlDbType.Date)
                    .Value =
                        validatedProjectStartDate;

                restoreProjectCommand.Parameters.Add(
                    parameterName: "@TimeZoneId",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 200)
                    .Value = validatedTimeZoneId;

                restoreProjectCommand.Parameters.Add(
                    parameterName: "@DefaultReportOutputPath",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 1000)
                    .Value = validatedOutputPath;

                int restoredRows =
                    await restoreProjectCommand.ExecuteNonQueryAsync();

                if (restoredRows != 1)
                {
                    throw new InvalidOperationException(
                        $"Project '{validatedProjectName}' could not be restored because " +
                        "its database state changed before the operation completed.");
                }

                #endregion


                #region Report Successful Restoration

                txtProjectManagementStatus.Text =
                    $"Project '{validatedProjectName}' restored successfully.";

                return true;

                #endregion
            }

            #endregion


            #region Insert New Project

            const string insertProjectSql = """
        INSERT INTO [dbo].[Project]
        (
            [ProjectName],
            [ProjectStartDate],
            [TimeZoneId],
            [DefaultReportOutputPath],
            [IsDeleted]
        )
        VALUES
        (
            @ProjectName,
            @ProjectStartDate,
            @TimeZoneId,
            @DefaultReportOutputPath,
            0
        );
        """;

            await using SqlCommand insertProjectCommand =
                new(
                    cmdText: insertProjectSql,
                    connection: databaseConnection);

            insertProjectCommand.Parameters.AddWithValue(
                parameterName: "@ProjectName",
                value: validatedProjectName);

            insertProjectCommand.Parameters.Add(
                parameterName: "@ProjectStartDate",
                sqlDbType: System.Data.SqlDbType.Date)
                .Value =
                    validatedProjectStartDate;

            insertProjectCommand.Parameters.Add(
                parameterName: "@TimeZoneId",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 200)
                .Value = validatedTimeZoneId;

            insertProjectCommand.Parameters.Add(
                parameterName: "@DefaultReportOutputPath",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 1000)
                .Value = validatedOutputPath;

            int insertedRows =
                await insertProjectCommand.ExecuteNonQueryAsync();

            if (insertedRows != 1)
            {
                throw new InvalidOperationException(
                    $"Project '{validatedProjectName}' was not inserted.");
            }

            #endregion


            #region Report Successful Addition

            txtProjectManagementStatus.Text =
                $"Project '{validatedProjectName}' added successfully.";

            return true;

            #endregion
        }

        #endregion

        #region Active Project Selection

        private async void ActiveProjectRadioButton_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Event Source

            if (sender is not RadioButton activeProjectRadioButton)
            {
                throw new InvalidOperationException(
                    "Active project selection was raised by an invalid control.");
            }

            if (activeProjectRadioButton.DataContext
                is not ProjectConfigurationItem selectedProject)
            {
                throw new InvalidOperationException(
                    "Active project selection does not contain a valid project.");
            }

            #endregion


            #region Prevent Deleted Project Selection

            if (selectedProject.IsDeleted)
            {
                ApplyActiveProjectStateToLoadedProjects();

                txtProjectManagementStatus.Text =
                    "A deleted project cannot be selected as the active project.";

                return;
            }

            #endregion


            #region Handle Existing Active Selection

            if (_activeProjectId.HasValue &&
                _activeProjectId.Value == selectedProject.Project_ID)
            {
                selectedProject.IsActive =
                    true;

                txtProjectManagementStatus.Text =
                    $"'{selectedProject.ProjectName}' is already the active project.";

                return;
            }

            #endregion


            #region Preserve Existing Project Lock

            // Do not release the existing Shared lock yet.
            //
            // The current project must remain fully operational unless the complete
            // transition to the newly selected project succeeds.

            SqlConnection? previousProjectLockConnection =
                _activeProjectLockConnection;

            SqlConnection? newProjectLockConnection =
                null;

            #endregion








            try
            {
                #region Acquire New Project Shared Lock

                newProjectLockConnection =
                    await OpenSharedProjectLockConnectionAsync(
                        projectId: selectedProject.Project_ID);

                #endregion


                #region Validate New Project Against Database

                const string validateProjectSql = """
            SELECT
                [ProjectName],
                [IsDeleted]
            FROM [dbo].[Project]
            WHERE [Project_ID] = @Project_ID;
            """;

                string databaseProjectName;
                bool databaseProjectIsDeleted;

                await using (SqlCommand validateProjectCommand =
                    new(
                        cmdText: validateProjectSql,
                        connection: newProjectLockConnection))
                {
                    validateProjectCommand.Parameters.AddWithValue(
                        parameterName: "@Project_ID",
                        value: selectedProject.Project_ID);

                    await using SqlDataReader reader =
                        await validateProjectCommand.ExecuteReaderAsync();

                    if (!await reader.ReadAsync())
                    {
                        throw new InvalidOperationException(
                            $"Project_ID {selectedProject.Project_ID} no longer exists.");
                    }

                    databaseProjectName =
                        reader.GetString(
                            i: 0);

                    databaseProjectIsDeleted =
                        reader.GetBoolean(
                            i: 1);
                }

                if (databaseProjectIsDeleted)
                {
                    throw new InvalidOperationException(
                        $"Project '{databaseProjectName}' has been deleted and " +
                        "cannot be selected as the active project.");
                }

                #endregion


                #region Persist New Startup Default

                // Persist only after the new Shared lock has been acquired and
                // the SQL project record has been validated.
                //
                // This changes the default for FUTURE application instances only.
                // Existing application instances retain their own process-local
                // active-project state.

                SaveActiveProjectToRegistry(
                    projectId: selectedProject.Project_ID,
                    projectName: databaseProjectName);

                #endregion


                #region Clear Existing DataGrid Active Flags

                foreach (ProjectConfigurationItem projectItem in _projectItems)
                {
                    projectItem.IsActive =
                        false;
                }

                #endregion


                #region Establish New Process-Local Active Project

                // SQL ProjectName is authoritative in case another application
                // instance has renamed the project since this DataGrid was loaded.

                selectedProject.ProjectName =
                    databaseProjectName;

                selectedProject.OriginalProjectName =
                    databaseProjectName;

                selectedProject.IsActive =
                    true;

                _activeProjectId =
                    selectedProject.Project_ID;

                _activeProjectName =
                    databaseProjectName;

                txtActiveProject.Text =
                    databaseProjectName;

                UpdateConfigurationWorkflowTabAvailability();

                #endregion


                #region Transfer Active Project Lock Ownership

                // The newly opened SQL connection now becomes the dedicated Shared
                // lock connection for this application instance.

                _activeProjectLockConnection =
                    newProjectLockConnection;

                // Clear the local reference so the exception cleanup code cannot
                // dispose the connection now owned by the application instance.

                newProjectLockConnection =
                    null;

                #endregion

                #region Report Successful Project Change

                // Any loaded import source belongs to the project that was active
                // when that import was started.
                //
                // Changing the active project invalidates all existing import state.

                ResetReferenceImportState(
                    statusMessage:
                        $"Active project changed to '{_activeProjectName}'. " +
                        "Select a CSV for this project.");

                ResetGeotechImportState(
                    statusMessage:
                        $"Active project changed to '{_activeProjectName}'. " +
                        "Select a Geotech CSV for this project.");

                ResetPrismPairImportState(
                    statusMessage:
                        $"Active project changed to '{_activeProjectName}'. " +
                        "Select a Track Geometry workbook for this project.");

                ResetPrismArrayConfigurationState();

                if (tabPrismArrays.IsSelected)
                {
                    try
                    {
                        await InitialisePrismArrayConfigurationForActiveProjectAsync();

                        UpdatePrismArrayAssignmentDisplay();

                        txtPrismArrayStatus.Text =
                            $"{_availablePrismArrayPoints.Count} point(s) available; " +
                            $"{_committedPrismArrays.Count} array(s) committed.";
                    }
                    catch (Exception arrayEx)
                    {
                        txtPrismArrayStatus.Text =
                            $"Unable to load arrays: {arrayEx.Message}";
                    }
                }

                try
                {
                    await EnsureChartTypeCatalogueAsync();

                    await EnsureChartTemplateTablesAsync();

                    await RefreshExistingChartsAsync();

                    await RefreshChartTemplatesAsync();
                }
                catch (Exception chartEx)
                {
                    txtChartStatus.Text =
                        $"Unable to load chart list: {chartEx.Message}";
                }

                txtProjectManagementStatus.Text =
                    $"Active project set to '{_activeProjectName}'.";

                #endregion


            }
            catch (Exception ex)
            {
                #region Release Incomplete New Project Lock

                if (newProjectLockConnection is not null)
                {
                    await newProjectLockConnection.DisposeAsync();

                    newProjectLockConnection =
                        null;
                }

                #endregion


                #region Restore Existing DataGrid State

                // The process-local active-project fields and existing Shared lock
                // were deliberately not changed until the new project had been
                // successfully acquired and validated.

                ApplyActiveProjectStateToLoadedProjects();

                #endregion


                #region Report Selection Failure

                txtProjectManagementStatus.Text =
                    $"Unable to set active project: {ex.Message}";

                #endregion

                return;
            }


            #region Release Previous Project Shared Lock

            // The new project is now fully established.
            //
            // Only now is it safe to release the previous project's Shared lock.

            if (previousProjectLockConnection is not null &&
                !ReferenceEquals(
                    previousProjectLockConnection,
                    _activeProjectLockConnection))
            {
                try
                {
                    await previousProjectLockConnection.DisposeAsync();
                }
                catch (Exception ex)
                {
                    // The active-project change itself has succeeded.
                    //
                    // Report a lock-release warning rather than reverting the
                    // successfully established new project.

                    txtProjectManagementStatus.Text =
                        $"Active project set to '{_activeProjectName}', " +
                        $"but the previous project lock reported: {ex.Message}";
                }
            }

            #endregion
        }

        #endregion


        #region Project Rename

        private async void ProjectNameTextBox_LostFocus(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Event Source

            if (sender is not TextBox projectNameTextBox)
            {
                throw new InvalidOperationException(
                    "Project-name edit was raised by an invalid control.");
            }

            if (projectNameTextBox.DataContext
                is not ProjectConfigurationItem selectedProject)
            {
                throw new InvalidOperationException(
                    "Project-name edit does not contain a valid project.");
            }

            #endregion


            #region Read Edited Project Name

            string newProjectName =
                projectNameTextBox.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Project-name edit control returned null.");

            string originalProjectName =
                selectedProject.OriginalProjectName?.Trim()
                ?? throw new InvalidOperationException(
                    "Original project name is not available.");

            #endregion


            #region Handle Unchanged Project Name

            if (string.Equals(
                a: newProjectName,
                b: originalProjectName,
                comparisonType: StringComparison.Ordinal))
            {
                selectedProject.ProjectName =
                    originalProjectName;

                return;
            }

            #endregion


            #region Validate Edited Project Name

            if (string.IsNullOrWhiteSpace(newProjectName))
            {
                selectedProject.ProjectName =
                    originalProjectName;

                projectNameTextBox.Text =
                    originalProjectName;

                txtProjectManagementStatus.Text =
                    "Project name cannot be blank.";

                return;
            }

            if (newProjectName.Length > 200)
            {
                selectedProject.ProjectName =
                    originalProjectName;

                projectNameTextBox.Text =
                    originalProjectName;

                txtProjectManagementStatus.Text =
                    "Project name cannot exceed 200 characters.";

                return;
            }

            #endregion


            try
            {
                #region Rename Project In Database

                await RenameProjectAsync(
                    projectId: selectedProject.Project_ID,
                    originalProjectName: originalProjectName,
                    newProjectName: newProjectName);

                #endregion


                #region Update Local Project Item

                selectedProject.ProjectName =
                    newProjectName;

                selectedProject.OriginalProjectName =
                    newProjectName;

                #endregion


                #region Update Active Project Runtime State

                if (_activeProjectId.HasValue &&
                    _activeProjectId.Value == selectedProject.Project_ID)
                {
                    // Project_ID remains the authoritative operational identity.
                    //
                    // Only the process-local display name changes.

                    _activeProjectName =
                        newProjectName;

                    if (_prismArrayProjectId.HasValue &&
                        _prismArrayProjectId.Value == selectedProject.Project_ID)
                    {
                        _prismArrayProjectName =
                            newProjectName;
                    }

                    txtActiveProject.Text =
                        newProjectName;

                    #region Update Persisted Startup Name Safely

                    UpdatePersistedActiveProjectNameIfCurrent(
                        projectId: selectedProject.Project_ID,
                        projectName: newProjectName);

                    #endregion
                }

                #endregion


                #region Report Successful Rename

                txtProjectManagementStatus.Text =
                    $"Project renamed to '{newProjectName}'.";

                #endregion
            }

            #region Handle Duplicate Project Name

            catch (SqlException ex)
                when (ex.Number == 2601 ||
                      ex.Number == 2627)
            {
                await LoadProjectsAsync();

                txtProjectManagementStatus.Text =
                    $"Project name '{newProjectName}' already exists.";
            }

            #endregion


            #region Handle Rename Errors

            catch (Exception ex)
            {
                // SQL remains authoritative.
                // Reload the complete list to remove any uncommitted UI value.

                await LoadProjectsAsync();

                txtProjectManagementStatus.Text =
                    $"Unable to rename project: {ex.Message}";
            }

            #endregion
        }


        private async Task RenameProjectAsync(
            int projectId,
            string originalProjectName,
            string newProjectName)
        {
            #region Validate Project ID

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(projectId),
                    message: "Project_ID must be greater than zero.");
            }

            #endregion


            #region Validate Original Project Name

            string validatedOriginalProjectName =
                originalProjectName?.Trim()
                ?? throw new ArgumentNullException(
                    paramName: nameof(originalProjectName));

            if (string.IsNullOrWhiteSpace(validatedOriginalProjectName))
            {
                throw new ArgumentException(
                    message: "Original project name cannot be empty.",
                    paramName: nameof(originalProjectName));
            }

            #endregion


            #region Validate New Project Name

            string validatedNewProjectName =
                newProjectName?.Trim()
                ?? throw new ArgumentNullException(
                    paramName: nameof(newProjectName));

            if (string.IsNullOrWhiteSpace(validatedNewProjectName))
            {
                throw new ArgumentException(
                    message: "New project name cannot be empty.",
                    paramName: nameof(newProjectName));
            }

            if (validatedNewProjectName.Length > 200)
            {
                throw new ArgumentException(
                    message: "New project name cannot exceed 200 characters.",
                    paramName: nameof(newProjectName));
            }

            #endregion


            #region Resolve Database Connection String

            string databaseConnectionString =
                GetTrackGeometryConnectionString();

            #endregion


            #region Define Rename Command

            const string renameProjectSql = """
        UPDATE [dbo].[Project]
        SET
            [ProjectName] = @NewProjectName
        WHERE
            [Project_ID] = @Project_ID
            AND [ProjectName] = @OriginalProjectName;
        """;

            #endregion


            #region Rename Project

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionString);

            await databaseConnection.OpenAsync();

            await using SqlCommand renameProjectCommand =
                new(
                    cmdText: renameProjectSql,
                    connection: databaseConnection);

            renameProjectCommand.Parameters.AddWithValue(
                parameterName: "@Project_ID",
                value: projectId);

            renameProjectCommand.Parameters.AddWithValue(
                parameterName: "@OriginalProjectName",
                value: validatedOriginalProjectName);

            renameProjectCommand.Parameters.AddWithValue(
                parameterName: "@NewProjectName",
                value: validatedNewProjectName);

            int affectedRows =
                await renameProjectCommand.ExecuteNonQueryAsync();

            if (affectedRows != 1)
            {
                throw new InvalidOperationException(
                    $"Project_ID {projectId} could not be renamed because its " +
                    "database state changed after this Project Management window " +
                    "was loaded.");
            }

            #endregion
        }


        private static void UpdatePersistedActiveProjectNameIfCurrent(
            int projectId,
            string projectName)
        {
            #region Validate Parameters

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(projectId),
                    message: "Project_ID must be greater than zero.");
            }

            string validatedProjectName =
                projectName?.Trim()
                ?? throw new ArgumentNullException(
                    paramName: nameof(projectName));

            if (string.IsNullOrWhiteSpace(validatedProjectName))
            {
                throw new ArgumentException(
                    message: "Project name cannot be empty.",
                    paramName: nameof(projectName));
            }

            #endregion


            #region Open Persisted Startup Configuration

            using RegistryKey? registryKey =
                Registry.CurrentUser.OpenSubKey(
                    name: RegistryPath,
                    writable: true);

            if (registryKey is null)
            {
                return;
            }

            #endregion


            #region Verify Persisted Startup Project

            object? persistedProjectIdValue =
                registryKey.GetValue(
                    name: RegistryActiveProjectId,
                    defaultValue: null);

            if (persistedProjectIdValue is not int persistedProjectId ||
                persistedProjectId != projectId)
            {
                // Another running instance has changed the persisted startup
                // project since this process started.
                //
                // Do not overwrite that newer startup default.

                return;
            }

            #endregion


            #region Update Persisted Project Name

            registryKey.SetValue(
                name: RegistryActiveProjectName,
                value: validatedProjectName,
                valueKind: RegistryValueKind.String);

            #endregion
        }

        #endregion




        #region Project Soft Delete And Restore

        private async void DeletedProjectCheckBox_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Event Source

            if (sender is not CheckBox deletedProjectCheckBox)
            {
                throw new InvalidOperationException(
                    "Project deleted-state change was raised by an invalid control.");
            }

            if (deletedProjectCheckBox.DataContext
                is not ProjectConfigurationItem selectedProject)
            {
                throw new InvalidOperationException(
                    "Project deleted-state change does not contain a valid project.");
            }

            #endregion


            #region Determine Requested State

            bool requestedDeletedState =
                deletedProjectCheckBox.IsChecked == true;

            #endregion


            #region Prevent Active Project Deletion

            if (requestedDeletedState &&
                (_activeProjectId == selectedProject.Project_ID ||
                 selectedProject.IsActive))
            {
                // Restore the visual state because an active project
                // is never permitted to become deleted.

                selectedProject.IsDeleted =
                    false;

                deletedProjectCheckBox.IsChecked =
                    false;

                txtProjectManagementStatus.Text =
                    $"Project '{selectedProject.ProjectName}' is active and cannot be deleted.";

                return;
            }

            #endregion


            #region Prepare User Interface

            deletedProjectCheckBox.IsEnabled =
                false;

            txtProjectManagementStatus.Text =
                requestedDeletedState
                    ? $"Deleting project '{selectedProject.ProjectName}'..."
                    : $"Restoring project '{selectedProject.ProjectName}'...";

            #endregion


            try
            {
                #region Update Project Deleted State

                // Persist the requested soft-delete or restore state to SQL Server.
                //
                // SetProjectDeletedStateAsync() also enforces the cross-instance rule:
                // deletion requires an Exclusive application lock and therefore cannot
                // proceed while the project is active in another running instance.

                await SetProjectDeletedStateAsync(
                    projectId: selectedProject.Project_ID,
                    isDeleted: requestedDeletedState);

                #endregion


                #region Reload Projects From Database

                // SQL is authoritative.
                //
                // Reload the complete project list after every change so this
                // running instance does not rely on stale DataGrid state.

                await LoadProjectsAsync();

                #endregion


                #region Report Successful Completion

                txtProjectManagementStatus.Text =
                    requestedDeletedState
                        ? $"Project '{selectedProject.ProjectName}' deleted."
                        : $"Project '{selectedProject.ProjectName}' restored.";

                #endregion
            }

            #region Handle Project State Errors

            catch (SqlException ex)
            {
                await LoadProjectsAsync();

                txtProjectManagementStatus.Text =
                    $"Unable to change project state: {ex.Message}";
            }
            catch (ArgumentException ex)
            {
                await LoadProjectsAsync();

                txtProjectManagementStatus.Text =
                    $"Invalid project state: {ex.Message}";
            }
            catch (InvalidOperationException ex)
            {
                await LoadProjectsAsync();

                txtProjectManagementStatus.Text =
                    $"Unable to change project state: {ex.Message}";
            }
            catch (Exception ex)
            {
                await LoadProjectsAsync();

                txtProjectManagementStatus.Text =
                    $"Unable to change project state: {ex.Message}";
            }

            #endregion
        }


        private async Task SetProjectDeletedStateAsync(
    int projectId,
    bool isDeleted)
        {
            #region Validate Project

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(projectId),
                    message: "Project_ID must be greater than zero.");
            }

            #endregion


            #region Protect Process-Local Active Project

            if (isDeleted &&
                _activeProjectId.HasValue &&
                _activeProjectId.Value == projectId)
            {
                throw new InvalidOperationException(
                    "The active project cannot be deleted.");
            }

            #endregion


            #region Handle Project Deletion

            if (isDeleted)
            {
                // -------------------------------------------------------------
                // CROSS-INSTANCE DELETE RULE
                //
                // Deletion requires an Exclusive application lock.
                //
                // Any running instance that has this project active holds a
                // Shared lock on exactly the same SQL application-lock resource.
                //
                // Shared + Exclusive are incompatible. Therefore this operation
                // cannot proceed while any application instance is actively
                // using this project.
                // -------------------------------------------------------------

                #region Acquire Exclusive Project Lock

                await using SqlConnection exclusiveLockConnection =
                    await OpenExclusiveProjectLockConnectionAsync(
                        projectId: projectId);

                #endregion


                #region Revalidate Project Under Exclusive Lock

                const string validateProjectSql = """
            SELECT
                [ProjectName],
                [IsDeleted]
            FROM [dbo].[Project]
            WHERE [Project_ID] = @Project_ID;
            """;

                string projectName;
                bool projectAlreadyDeleted;

                await using (SqlCommand validateProjectCommand =
                    new(
                        cmdText: validateProjectSql,
                        connection: exclusiveLockConnection))
                {
                    validateProjectCommand.Parameters.AddWithValue(
                        parameterName: "@Project_ID",
                        value: projectId);

                    await using SqlDataReader reader =
                        await validateProjectCommand.ExecuteReaderAsync();

                    if (!await reader.ReadAsync())
                    {
                        throw new InvalidOperationException(
                            $"Project_ID {projectId} no longer exists.");
                    }

                    projectName =
                        reader.GetString(
                            i: 0);

                    projectAlreadyDeleted =
                        reader.GetBoolean(
                            i: 1);
                }

                #endregion


                #region Handle Already Deleted Project

                if (projectAlreadyDeleted)
                {
                    throw new InvalidOperationException(
                        $"Project '{projectName}' is already deleted.");
                }

                #endregion


                #region Delete Project Under Exclusive Lock

                const string deleteProjectSql = """
            UPDATE [dbo].[Project]
            SET
                [IsDeleted] = 1
            WHERE
                [Project_ID] = @Project_ID
                AND [IsDeleted] = 0;
            """;

                await using SqlCommand deleteProjectCommand =
                    new(
                        cmdText: deleteProjectSql,
                        connection: exclusiveLockConnection);

                deleteProjectCommand.Parameters.AddWithValue(
                    parameterName: "@Project_ID",
                    value: projectId);

                int affectedRows =
                    await deleteProjectCommand.ExecuteNonQueryAsync();

                if (affectedRows != 1)
                {
                    throw new InvalidOperationException(
                        $"Project '{projectName}' could not be deleted because " +
                        "its database state changed before the operation completed.");
                }

                #endregion


                #region Complete Exclusive Delete Operation

                // ExecuteNonQueryAsync() has completed the SQL UPDATE before this
                // point.
                //
                // The Exclusive lock remains held until exclusiveLockConnection
                // is disposed when this scope exits.
                //
                // Only after that disposal may another instance acquire a Shared
                // lock. Such an instance will then validate IsDeleted = 1 and
                // reject the project as an active-project selection.

                return;

                #endregion
            }

            #endregion


            #region Handle Project Restoration

            // A deleted project cannot legitimately be active, therefore restoring
            // it does not require the Exclusive active-project protection used
            // during deletion.

            #region Resolve Database Connection String

            string databaseConnectionString =
                GetTrackGeometryConnectionString();

            #endregion


            #region Restore Project

            const string restoreProjectSql = """
        UPDATE [dbo].[Project]
        SET
            [IsDeleted] = 0
        WHERE
            [Project_ID] = @Project_ID
            AND [IsDeleted] = 1;
        """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionString);

            await databaseConnection.OpenAsync();

            await using SqlCommand restoreProjectCommand =
                new(
                    cmdText: restoreProjectSql,
                    connection: databaseConnection);

            restoreProjectCommand.Parameters.AddWithValue(
                parameterName: "@Project_ID",
                value: projectId);

            int restoredRows =
                await restoreProjectCommand.ExecuteNonQueryAsync();

            if (restoredRows != 1)
            {
                throw new InvalidOperationException(
                    $"Project_ID {projectId} could not be restored. " +
                    "The project may already have been restored or its database " +
                    "state may have changed.");
            }

            #endregion

            #endregion
        }

        #endregion




        private async Task LoadProjectsAsync()
        {
            #region Resolve Database Connection String

            string databaseConnectionString =
                GetTrackGeometryConnectionString();

            #endregion


            #region Define Project Query

            const string loadProjectsSql = """
                SELECT
                    [Project_ID],
                    [ProjectName],
                    [ProjectStartDate],
                    [IsDeleted]
                FROM [dbo].[Project]
                ORDER BY
                    [ProjectName];
                """;

            #endregion


            #region Clear Existing Project List

            _projectItems.Clear();

            _projectStartDates.Clear();

            #endregion


            #region Load Projects From Database

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionString);

            await databaseConnection.OpenAsync();

            await using SqlCommand loadProjectsCommand =
                new(
                    cmdText: loadProjectsSql,
                    connection: databaseConnection);

            await using SqlDataReader reader =
                await loadProjectsCommand.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                int projectId =
                    reader.GetInt32(
                        i: 0);

                string projectName =
                    reader.GetString(
                        i: 1);

                DateTime projectStartDate =
                    reader.GetDateTime(
                        i: 2)
                    .Date;

                bool isDeleted =
                    reader.GetBoolean(
                        i: 3);

                _projectStartDates[projectId] =
                    projectStartDate;

                ProjectConfigurationItem projectItem =
                    new()
                    {
                        Project_ID = projectId,
                        ProjectName = projectName,
                        OriginalProjectName = projectName,
                        IsDeleted = isDeleted,
                        IsActive = false
                    };

                _projectItems.Add(
                    item: projectItem);
            }

            #endregion


            #region Apply Process-Local Active Project

            ApplyActiveProjectStateToLoadedProjects();

            UpdateAddProjectButtonState();

            #endregion
        }


        private void ApplyActiveProjectStateToLoadedProjects()
        {
            #region Reset DataGrid Active State

            foreach (ProjectConfigurationItem projectItem in _projectItems)
            {
                projectItem.IsActive =
                    false;
            }

            #endregion


            #region Handle No Active Project

            if (!_activeProjectId.HasValue)
            {
                _activeProjectName =
                    string.Empty;

                txtActiveProject.Text =
                    "No active project";

                UpdateConfigurationWorkflowTabAvailability();

                return;
            }

            #endregion


            #region Locate Active Project

            ProjectConfigurationItem? activeProjectItem =
                null;

            foreach (ProjectConfigurationItem projectItem in _projectItems)
            {
                if (projectItem.Project_ID == _activeProjectId.Value)
                {
                    activeProjectItem =
                        projectItem;

                    break;
                }
            }

            #endregion


            #region Validate Active Project

            if (activeProjectItem is null ||
                activeProjectItem.IsDeleted)
            {
                // The startup default no longer represents an available
                // project in the current database.
                //
                // Only this process's runtime state is cleared here.
                //
                // The Registry is not modified during normal reconciliation
                // because another running instance may have changed the
                // persisted startup default after this process started.

                _activeProjectId =
                    null;

                _activeProjectName =
                    string.Empty;

                txtActiveProject.Text =
                    "No active project";

                UpdateConfigurationWorkflowTabAvailability();

                return;
            }

            #endregion


            #region Apply Active Project To This Process

            activeProjectItem.IsActive =
                true;

            // Project_ID is authoritative.
            // Refresh the process-local name from SQL in case the project
            // name has been changed since this process started.

            _activeProjectName =
                activeProjectItem.ProjectName;

            txtActiveProject.Text =
                _activeProjectName;

            UpdateConfigurationWorkflowTabAvailability();

            #endregion
        }
        #endregion


        #region Chart Configuration UI

        private void InitialiseChartConfigurationUi()
        {
            #region Populate Chart Type Catalogue

            _chartTypes.Clear();

            AddChartTypeUiItem(
                key: "TiltDegrees",
                displayName: "Tilt - Degrees",
                sourceTable: "TiltEpochs",
                entityKind: "Sensor",
                verticalAxisTitle: "Rotation",
                unit: "decimal degrees",
                dataElements: new[] { "TiltA", "TiltB", "TiltC" });

            AddChartTypeUiItem(
                key: "TiltMmPerM",
                displayName: "Tilt - mm/m",
                sourceTable: "TiltEpochs",
                entityKind: "Sensor",
                verticalAxisTitle: "Displacement",
                unit: "mm/m",
                dataElements: new[] { "TiltA", "TiltB", "TiltC" },
                isFullyDefined: false);

            AddChartTypeUiItem(
                key: "PrismTiltMmPerM",
                displayName: "Prism Tilt - mm/m",
                sourceTable: "PrismTiltEpochs",
                entityKind: "PrismArray",
                verticalAxisTitle: "Displacement",
                unit: "mm/m",
                dataElements: new[] { "TiltX", "TiltY" });

            AddChartTypeUiItem(
                key: "CrackMeter",
                displayName: "Crack Meter",
                sourceTable: "PrismCrackGaugeEpochs",
                entityKind: "PrismArray",
                verticalAxisTitle: "Displacement",
                unit: "mm",
                dataElements: new[] { "d2D", "d3D", "dH" });

            AddChartTypeUiItem(
                key: "Displacement",
                displayName: "Displacement",
                sourceTable: "CoordinatesEpochs",
                entityKind: "Point",
                verticalAxisTitle: "Displacement",
                unit: "mm",
                dataElements: new[] { "dE", "dN", "d2D" });


            AddChartTypeUiItem(
                key: "StructuralArray",
                displayName: "Structural Array",
                sourceTable: "StructuralArrayEpochs",
                entityKind: "PrismArray",
                verticalAxisTitle: "Displacement",
                unit: "mm",
                dataElements: new[] { "dE", "dN", "dH", "d2D", "d3D" });

            AddChartTypeUiItem(
                key: "Cant",
                displayName: "Cant",
                sourceTable: "CantEpochs",
                entityKind: "PrismPair",
                verticalAxisTitle: "Cant",
                unit: "mm",
                dataElements: new[] { "Cant" });

            AddChartTypeUiItem(
                key: "ShortTwistMmPer3m",
                displayName: "Short Twist - mm/3m",
                sourceTable: "ShortTwistEpochs",
                entityKind: "PrismPair",
                verticalAxisTitle: "Twist / 3m",
                unit: "mm/3m",
                dataElements: new[] { "ShortTwist" });

            AddChartTypeUiItem(
                key: "ShortTwistRatio",
                displayName: "Short Twist - Ratio",
                sourceTable: "ShortTwistEpochs",
                entityKind: "PrismPair",
                verticalAxisTitle: "Twist / 3m",
                unit: "ratio",
                dataElements: new[] { "ShortTwistRatio" });

            AddChartTypeUiItem(
                key: "LongTwistMmPer15m",
                displayName: "Long Twist - mm/15m",
                sourceTable: "LongTwistEpochs",
                entityKind: "PrismPair",
                verticalAxisTitle: "Twist / 15m",
                unit: "mm/15m",
                dataElements: new[] { "LongTwist" });

            AddChartTypeUiItem(
                key: "LongTwistRatio",
                displayName: "Long Twist - Ratio",
                sourceTable: "LongTwistEpochs",
                entityKind: "PrismPair",
                verticalAxisTitle: "Twist / 15m",
                unit: "ratio",
                dataElements: new[] { "LongTwistRatio" });

            AddChartTypeUiItem(
                key: "Slew",
                displayName: "Slew",
                sourceTable: "SlewEpochs",
                entityKind: "Point",
                verticalAxisTitle: "Slew",
                unit: "mm",
                dataElements: new[] { "Slew" },
                isFullyDefined: false);

            AddChartTypeUiItem(
                key: "LevelDh",
                displayName: "Level - dH",
                sourceTable: "DhEpochs",
                entityKind: "Point",
                verticalAxisTitle: "dH",
                unit: "mm",
                dataElements: new[] { "dH" });

            AddChartTypeUiItem(
                key: "Top",
                displayName: "Top",
                sourceTable: "TopEpochs",
                entityKind: "Point",
                verticalAxisTitle: "dH",
                unit: "mm",
                dataElements: new[] { "Top" });

            AddChartTypeUiItem(
                key: "ToRLevel",
                displayName: "ToR Level",
                sourceTable: "ToREpochs",
                entityKind: "Point",
                verticalAxisTitle: "dH",
                unit: "mm",
                dataElements: new[] { "ToR - Reference ToR" });

            AddChartTypeUiItem(
                key: "ConvergenceArray",
                displayName: "Convergence Array",
                sourceTable: "TunnelConvergenceEpochs",
                entityKind: "PrismArray",
                verticalAxisTitle: "Displacement",
                unit: "mm",
                dataElements: new[]
                {
                    "dAB", "dAC", "dAD", "dAE", "dBC",
                    "dBD", "dBE", "dCD", "dCE", "dDE"
                });

            AddChartTypeUiItem(
                key: "CrownData",
                displayName: "Crown Data",
                sourceTable: "StructuralArrayEpochs",
                entityKind: "PrismArray",
                verticalAxisTitle: "Displacement",
                unit: "mm",
                dataElements: new[] { "dE", "dN", "dH", "d2D", "d3D" });

            List<ChartTypeUiItem> sortedChartTypes =
                _chartTypes
                    .OrderBy(
                        keySelector: item => item.DisplayName,
                        comparer: StringComparer.OrdinalIgnoreCase)
                    .ToList();

            cmbChartType.ItemsSource =
                sortedChartTypes;

            if (sortedChartTypes.Count > 0)
            {
                SelectDefaultChartType();
            }

            #endregion


            #region Initialise Chart Dates

            dpChartEndDate.SelectedDate =
                DateTime.Today;

            dpChartStartDate.SelectedDate =
                DateTime.Today.AddDays(
                    value: -DefaultChartRecentDays);

            txtChartRecentDays.Text =
                DefaultChartRecentDays.ToString(
                    provider: CultureInfo.InvariantCulture);

            #endregion


            #region Initialise Report-Resolved Chart Configuration

            SelectChartStartDateMode(
                startDateMode: ChartStartDateModeReportStart);

            txtChartWidthMm.Text =
                DefaultChartWidthMm.ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartHeightMm.Text =
                DefaultChartHeightMm.ToString(
                    provider: CultureInfo.InvariantCulture);

            SelectChartResolutionDpi(
                resolutionDpi: DefaultChartResolutionDpi);

            UpdateChartPixelDimensions();

            #endregion


            #region Initialise Chart Series

            dgChartSeries.ItemsSource =
                _chartSeries;

            #endregion


            #region Initialise Y Axis Controls

            UpdateChartYAxisControlState();

            #endregion
        }


        #region Chart Default Values

        private void SelectDefaultChartType()
        {
            foreach (object item
                in cmbChartType.Items)
            {
                if (item is ChartTypeUiItem chartType &&
                    string.Equals(
                        a: chartType.Key,
                        b: "Displacement",
                        comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    cmbChartType.SelectedItem =
                        chartType;

                    return;
                }
            }

            if (cmbChartType.Items.Count > 0)
            {
                cmbChartType.SelectedIndex =
                    0;
            }
        }


        private static bool ChartUnitUsesStandardTriggerDefaults(
            string unit)
        {
            if (string.IsNullOrWhiteSpace(unit))
            {
                return false;
            }

            return unit.Contains(
                       value: "mm",
                       comparisonType: StringComparison.OrdinalIgnoreCase)
                   ||
                   string.Equals(
                       a: unit.Trim(),
                       b: "decimal degrees",
                       comparisonType: StringComparison.OrdinalIgnoreCase);
        }


        private void ApplyChartTypeDefaults(
            ChartTypeUiItem chartType)
        {
            #region Apply Standard Trigger Defaults

            if (ChartUnitUsesStandardTriggerDefaults(
                unit: chartType.Unit))
            {
                txtChartGreenTrigger.Text =
                    "2";

                txtChartAmberTrigger.Text =
                    "4";

                txtChartRedTrigger.Text =
                    "6";

                chkChartAutomaticYAxis.IsChecked =
                    true;

                ApplyDefaultVerticalAxisFromRedTrigger();
            }

            #endregion


            #region Apply Standard Presentation Defaults

            txtChartYAxisTitle.Text =
                chartType.VerticalAxisTitle;

            txtChartWidthMm.Text =
                DefaultChartWidthMm.ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartHeightMm.Text =
                DefaultChartHeightMm.ToString(
                    provider: CultureInfo.InvariantCulture);

            SelectChartResolutionDpi(
                resolutionDpi: DefaultChartResolutionDpi);

            UpdateChartPixelDimensions();

            txtChartDefaultLineWidth.Text =
                "2";

            txtChartDefaultMarkerSize.Text =
                "4";

            chkChartGridLines.IsChecked =
                true;

            chkChartShowLegend.IsChecked =
                true;

            cmbChartLegendPosition.SelectedIndex =
                0;

            #endregion
        }


        private void ApplyDefaultVerticalAxisFromRedTrigger()
        {
            if (!decimal.TryParse(
                s: txtChartRedTrigger.Text,
                style: NumberStyles.Float,
                provider: CultureInfo.InvariantCulture,
                result: out decimal redTrigger) ||
                redTrigger <= 0)
            {
                return;
            }

            decimal rawExtent =
                redTrigger * 1.5m;

            decimal roundedExtent =
                Math.Ceiling(
                    d: rawExtent / 5m) *
                5m;

            if (roundedExtent < 5m)
            {
                roundedExtent =
                    5m;
            }

            txtChartYAxisMinimum.Text =
                (-roundedExtent).ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartYAxisMaximum.Text =
                roundedExtent.ToString(
                    provider: CultureInfo.InvariantCulture);
        }


        private void txtChartRedTrigger_LostFocus(
            object sender,
            RoutedEventArgs e)
        {
            if (chkChartAutomaticYAxis.IsChecked == true)
            {
                ApplyDefaultVerticalAxisFromRedTrigger();
            }
        }

        #endregion


        private void AddChartTypeUiItem(
            string key,
            string displayName,
            string sourceTable,
            string entityKind,
            string verticalAxisTitle,
            string unit,
            IReadOnlyList<string> dataElements,
            bool isFullyDefined = true)
        {
            _chartTypes.Add(
                item:
                    new ChartTypeUiItem
                    {
                        Key = key,
                        DisplayName = displayName,
                        SourceTable = sourceTable,
                        EntityKind = entityKind,
                        VerticalAxisTitle = verticalAxisTitle,
                        Unit = unit,
                        DataElements = dataElements,
                        IsFullyDefined = isFullyDefined
                    });
        }


        private async void cmbChartType_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Apply Selected Chart Type

            if (cmbChartType.SelectedItem
                is not ChartTypeUiItem selectedType)
            {
                txtChartDataSource.Text =
                    string.Empty;

                txtChartUnit.Text =
                    string.Empty;

                txtChartTriggerUnit.Text =
                    string.Empty;

                cmbChartDataElement.ItemsSource =
                    null;

                cmbChartSeriesEntity.ItemsSource =
                    null;

                return;
            }

            txtChartDataSource.Text =
                $"{selectedType.SourceTable} ({selectedType.EntityKind})";

            txtChartUnit.Text =
                selectedType.Unit;

            txtChartTriggerUnit.Text =
                selectedType.Unit;

            ApplyChartTypeDefaults(
                chartType: selectedType);

            cmbChartDataElement.ItemsSource =
                selectedType.DataElements;

            if (selectedType.DataElements.Count > 0)
            {
                cmbChartDataElement.SelectedIndex =
                    0;
            }

            try
            {
                await LoadChartEntitiesAsync(
                    chartType: selectedType);
            }
            catch (Exception ex)
            {
                cmbChartSeriesEntity.ItemsSource =
                    null;

                txtChartStatus.Text =
                    $"Unable to load chart entities: {ex.Message}";

                return;
            }

            if (!selectedType.IsFullyDefined)
            {
                txtChartStatus.Text =
                    $"'{selectedType.DisplayName}' is available for configuration, " +
                    "but its final engineering conversion/label rule remains to be confirmed.";
            }
            else
            {
                txtChartStatus.Text =
                    $"Chart type '{selectedType.DisplayName}' selected.";
            }

            #endregion
        }


        private async Task LoadChartEntitiesAsync(
            ChartTypeUiItem chartType)
        {
            cmbChartSeriesEntity.ItemsSource =
                null;

            if (!_activeProjectId.HasValue)
            {
                return;
            }

            string sql = chartType.EntityKind switch
            {
                "Point" => """
                    SELECT [PointName_ID], [PointName]
                    FROM [dbo].[PointName]
                    WHERE [Project_ID] = @Project_ID AND [IsDeleted] = 0
                    ORDER BY [PointName];
                    """,
                "Sensor" => """
                    SELECT [SensorID], [SensorName]
                    FROM [dbo].[GeotecSensors]
                    WHERE [Project_ID] = @Project_ID AND [IsDeleted] = 0
                    ORDER BY [SensorName];
                    """,
                "Track" => """
                    SELECT [Track_ID], [TrackName]
                    FROM [dbo].[Track]
                    WHERE [Project_ID] = @Project_ID AND [IsDeleted] = 0
                    ORDER BY [TrackName];
                    """,
                "PrismArray" => """
                    SELECT [Array_ID], [ArrayName]
                    FROM [dbo].[PrismArray]
                    WHERE [Project_ID] = @Project_ID AND [IsDeleted] = 0
                    ORDER BY [ArrayName];
                    """,
                "PrismPair" => """
                    SELECT
                        PP.[PrismPair_ID],
                        CONCAT(T.[TrackName], N' | ', L.[PointName], N' - ', R.[PointName])
                    FROM [dbo].[PrismPairs] AS PP
                    INNER JOIN [dbo].[Track] AS T ON T.[Track_ID] = PP.[Track_ID]
                    INNER JOIN [dbo].[PointName] AS L ON L.[PointName_ID] = PP.[Left_ID]
                    INNER JOIN [dbo].[PointName] AS R ON R.[PointName_ID] = PP.[Right_ID]
                    WHERE
                        T.[Project_ID] = @Project_ID
                        AND T.[IsDeleted] = 0
                        AND PP.[IsDeleted] = 0
                    ORDER BY T.[TrackName], PP.[PairOrder];
                    """,
                _ => throw new InvalidOperationException(
                    $"Unsupported chart entity kind '{chartType.EntityKind}'.")
            };

            List<ChartEntityUiItem> entities =
                new();

            await using SqlConnection databaseConnection =
                new(
                    connectionString: GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand command =
                new(
                    cmdText: sql,
                    connection: databaseConnection);

            command.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value = _activeProjectId.Value;

            await using SqlDataReader reader =
                await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                entities.Add(
                    item:
                        new ChartEntityUiItem
                        {
                            EntityId = reader.GetInt32(0),
                            DisplayName = reader.GetString(1)
                        });
            }

            HashSet<int> selectedEntityIds =
                _chartSeries
                    .Where(
                        predicate: series => series.EntityId > 0)
                    .Select(
                        selector: series => series.EntityId)
                    .ToHashSet();

            entities.RemoveAll(
                match: entity => selectedEntityIds.Contains(
                    item: entity.EntityId));

            cmbChartSeriesEntity.ItemsSource =
                entities;

            if (entities.Count > 0)
            {
                cmbChartSeriesEntity.SelectedIndex =
                    0;
            }
        }


        private void btnChartNew_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Reset New Chart Definition

            _loadedChartDefinitionId =
                null;

            _loadedChartNumber =
                null;

            _loadedChartTemplateId =
                null;

            cmbExistingChart.SelectedItem =
                null;

            cmbChartTemplate.SelectedItem =
                null;

            _loadedChartTemplateId =
                null;

            txtChartNumber.Text =
                "New";

            txtChartName.Clear();

            _chartSeries.Clear();

            SelectDefaultChartType();

            dpChartEndDate.SelectedDate =
                DateTime.Today;

            dpChartStartDate.SelectedDate =
                DateTime.Today.AddDays(
                    value: -DefaultChartRecentDays);

            txtChartRecentDays.Text =
                DefaultChartRecentDays.ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartGreenTrigger.Text =
                "2";

            txtChartAmberTrigger.Text =
                "4";

            txtChartRedTrigger.Text =
                "6";

            chkChartAutomaticYAxis.IsChecked =
                true;

            ApplyDefaultVerticalAxisFromRedTrigger();

            txtChartTitleOverride.Clear();

            chkChartShowLegend.IsChecked =
                true;

            cmbChartLegendPosition.SelectedIndex =
                0;

            txtChartPngWidth.Text =
                "3000";

            txtChartPngHeight.Text =
                "1000";

            txtChartDefaultLineWidth.Text =
                "2";

            txtChartDefaultMarkerSize.Text =
                "4";

            chkChartGridLines.IsChecked =
                true;

            txtChartStatus.Text =
                "New chart canvas ready with default values.";

            txtChartName.Focus();

            #endregion
        }


        private async void btnChartLoad_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Load Selected Existing Chart

            try
            {
                if (!_activeProjectId.HasValue)
                {
                    throw new InvalidOperationException(
                        "Select an active project.");
                }

                await RefreshExistingChartsAsync();

                if (cmbExistingChart.SelectedItem
                    is not ExistingChartUiItem selectedChart)
                {
                    txtChartStatus.Text =
                        "Select an existing chart.";

                    return;
                }

                await LoadChartDefinitionIntoEditorAsync(
                    chartDefinitionId: selectedChart.ChartDefinitionId);

                txtChartStatus.Text =
                    $"Loaded {selectedChart.DisplayText}.";
            }
            catch (Exception ex)
            {
                txtChartStatus.Text =
                    $"Unable to load chart: {ex.Message}";
            }

            #endregion
        }


        private void btnChartCopy_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Copy Current Chart Definition

            if (string.IsNullOrWhiteSpace(
                value: txtChartName.Text))
            {
                txtChartStatus.Text =
                    "Load or define a chart before creating a copy.";

                return;
            }

            // All current editor settings already represent the source chart.
            // Creating a copy therefore preserves all settings and only removes
            // the database identity and changes the name.

            _loadedChartDefinitionId =
                null;

            _loadedChartNumber =
                null;

            cmbExistingChart.SelectedItem =
                null;

            txtChartNumber.Text =
                "New";

            string sourceName =
                txtChartName.Text.Trim();

            txtChartName.Text =
                $"{sourceName}-Copy";

            txtChartStatus.Text =
                "Chart copied. All settings retained; edit the copy and Commit.";

            txtChartName.Focus();

            txtChartName.SelectAll();

            #endregion
        }


        private async void btnChartDelete_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Delete Target

            ExistingChartUiItem? selectedChart =
                cmbExistingChart.SelectedItem
                    as ExistingChartUiItem;

            int? chartDefinitionId =
                selectedChart?.ChartDefinitionId
                ?? _loadedChartDefinitionId;

            if (!chartDefinitionId.HasValue)
            {
                txtChartStatus.Text =
                    "Select or load a chart to delete.";

                return;
            }

            string chartName =
                selectedChart?.ChartName
                ?? txtChartName.Text.Trim();

            MessageBoxResult confirmation =
                MessageBox.Show(
                    owner: this,
                    messageBoxText:
                        $"Delete chart '{chartName}'?",
                    caption:
                        "Delete Chart",
                    button:
                        MessageBoxButton.YesNo,
                    icon:
                        MessageBoxImage.Warning,
                    defaultResult:
                        MessageBoxResult.No);

            if (confirmation != MessageBoxResult.Yes)
            {
                return;
            }

            #endregion


            #region Soft Delete Chart

            try
            {
                await SoftDeleteChartDefinitionAsync(
                    chartDefinitionId: chartDefinitionId.Value);

                await RefreshExistingChartsAsync();

                btnChartNew_Click(
                    sender: this,
                    e: new RoutedEventArgs());

                txtChartStatus.Text =
                    $"Chart '{chartName}' deleted.";
            }
            catch (Exception ex)
            {
                txtChartStatus.Text =
                    $"Unable to delete chart: {ex.Message}";
            }

            #endregion
        }


        private async void btnChartCommit_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Chart Definition For Save

            if (!TryValidateChartConfiguration(
                validationMessage: out string validationMessage))
            {
                txtChartStatus.Text =
                    validationMessage;

                return;
            }

            #endregion


            try
            {
                #region Ensure Catalogue Exists

                await EnsureChartTypeCatalogueAsync();

                #endregion


                #region Resolve Duplicate Chart Name

                string chartName =
                    txtChartName.Text.Trim();

                ExistingChartUiItem? duplicateChart =
                    await FindChartByNameAsync(
                        chartName: chartName,
                        excludeChartDefinitionId: _loadedChartDefinitionId);

                int? targetChartDefinitionId =
                    _loadedChartDefinitionId;

                int? targetChartNumber =
                    _loadedChartNumber;

                if (duplicateChart is not null)
                {
                    MessageBoxResult replaceResult =
                        MessageBox.Show(
                            owner: this,
                            messageBoxText:
                                "Duplicate chart found\n\n" +
                                $"Replace '{duplicateChart.ChartName}'?",
                            caption:
                                "Duplicate chart found",
                            button:
                                MessageBoxButton.YesNo,
                            icon:
                                MessageBoxImage.Warning,
                            defaultResult:
                                MessageBoxResult.No);

                    if (replaceResult != MessageBoxResult.Yes)
                    {
                        txtChartStatus.Text =
                            "Save rejected. Chart name unchanged.";

                        txtChartName.Focus();

                        txtChartName.SelectAll();

                        return;
                    }

                    // Replace means the existing duplicate chart is overwritten.
                    // If another chart was loaded and renamed to this duplicate
                    // name, retire the former loaded definition so the project
                    // still has one active chart with the requested name.

                    if (_loadedChartDefinitionId.HasValue &&
                        _loadedChartDefinitionId.Value != duplicateChart.ChartDefinitionId)
                    {
                        await SoftDeleteChartDefinitionAsync(
                            chartDefinitionId:
                                _loadedChartDefinitionId.Value);
                    }

                    targetChartDefinitionId =
                        duplicateChart.ChartDefinitionId;

                    targetChartNumber =
                        duplicateChart.ChartNumber;
                }

                #endregion


                #region Save Or Update Chart

                (int ChartDefinitionId, int ChartNumber) commitResult =
                    await SaveChartDefinitionAsync(
                        chartDefinitionId:
                            targetChartDefinitionId,
                        chartNumber:
                            targetChartNumber);

                _loadedChartDefinitionId =
                    commitResult.ChartDefinitionId;

                _loadedChartNumber =
                    commitResult.ChartNumber;

                txtChartNumber.Text =
                    $"Chart_{commitResult.ChartNumber:0000}";

                await RefreshExistingChartsAsync(
                    selectedChartDefinitionId:
                        commitResult.ChartDefinitionId);

                txtChartStatus.Text =
                    $"Chart '{chartName}' saved.";

                #endregion
            }
            catch (Exception ex)
            {
                txtChartStatus.Text =
                    $"Chart save failed: {ex.Message}";
            }
        }


        #region Chart Template Workflow

        private async void btnSaveChartTemplate_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Template

            if (!TryValidateChartCanvasConfiguration(
                validationMessage: out string validationMessage))
            {
                txtChartStatus.Text =
                    validationMessage;

                return;
            }

            string templateName =
                txtChartName.Text?.Trim()
                ?? string.Empty;

            if (string.IsNullOrWhiteSpace(templateName))
            {
                txtChartStatus.Text =
                    "Enter a Chart name to use as the Template name.";

                txtChartName.Focus();

                return;
            }

            #endregion


            try
            {
                #region Ensure Template Storage

                await EnsureChartTypeCatalogueAsync();

                await EnsureChartTemplateTablesAsync();

                #endregion


                #region Resolve Duplicate Template

                ChartTemplateUiItem? duplicateTemplate =
                    await FindChartTemplateByNameAsync(
                        templateName:
                            templateName);

                int? targetTemplateId =
                    duplicateTemplate?.ChartTemplateId;

                if (duplicateTemplate is not null)
                {
                    MessageBoxResult replaceResult =
                        MessageBox.Show(
                            owner: this,
                            messageBoxText:
                                "Duplicate template found\n\n" +
                                $"Replace '{duplicateTemplate.TemplateName}'?",
                            caption:
                                "Duplicate template found",
                            button:
                                MessageBoxButton.YesNo,
                            icon:
                                MessageBoxImage.Warning,
                            defaultResult:
                                MessageBoxResult.No);

                    if (replaceResult != MessageBoxResult.Yes)
                    {
                        txtChartStatus.Text =
                            "Template save rejected. Name unchanged.";

                        txtChartName.Focus();

                        txtChartName.SelectAll();

                        return;
                    }
                }

                #endregion


                #region Save Template

                int templateId =
                    await SaveChartTemplateAsync(
                        chartTemplateId:
                            targetTemplateId,
                        templateName:
                            templateName);

                _loadedChartTemplateId =
                    templateId;

                await RefreshChartTemplatesAsync(
                    selectedTemplateId:
                        templateId);

                txtChartStatus.Text =
                    $"Template '{templateName}' saved.";

                #endregion
            }
            catch (Exception ex)
            {
                txtChartStatus.Text =
                    $"Template save failed: {ex.Message}";
            }
        }


        private async void btnLoadChartTemplate_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Load Selected Template

            try
            {
                await EnsureChartTemplateTablesAsync();

                await RefreshChartTemplatesAsync();

                if (cmbChartTemplate.SelectedItem
                    is not ChartTemplateUiItem selectedTemplate)
                {
                    txtChartStatus.Text =
                        "Select a Template.";

                    return;
                }

                await LoadChartTemplateIntoEditorAsync(
                    chartTemplateId:
                        selectedTemplate.ChartTemplateId);

                txtChartStatus.Text =
                    $"Template '{selectedTemplate.TemplateName}' loaded. " +
                    "Rename the chart, select data entities and set report dates before Commit.";

                txtChartName.Focus();

                txtChartName.SelectAll();
            }
            catch (Exception ex)
            {
                txtChartStatus.Text =
                    $"Unable to load Template: {ex.Message}";
            }

            #endregion
        }


        private async Task EnsureChartTemplateTablesAsync()
        {
            #region Define Template Schema

            const string sql = """
                IF OBJECT_ID(N'dbo.ChartTemplate', N'U') IS NULL
                BEGIN
                    CREATE TABLE [dbo].[ChartTemplate]
                    (
                        [ChartTemplate_ID] int IDENTITY(1,1) NOT NULL,
                        [Project_ID] int NOT NULL,
                        [TemplateName] nvarchar(200) NOT NULL,
                        [ChartType_ID] int NOT NULL,
                        [DefaultDurationDays] int NOT NULL
                            CONSTRAINT [DF_ChartTemplate_DefaultDurationDays]
                            DEFAULT (14),
                        [GreenTrigger] decimal(18,6) NOT NULL,
                        [AmberTrigger] decimal(18,6) NOT NULL,
                        [RedTrigger] decimal(18,6) NOT NULL,
                        [UseAutomaticVerticalAxis] bit NOT NULL
                            CONSTRAINT [DF_ChartTemplate_UseAutomaticVerticalAxis]
                            DEFAULT (1),
                        [VerticalAxisMinimum] decimal(18,6) NOT NULL,
                        [VerticalAxisMaximum] decimal(18,6) NOT NULL,
                        [ChartTitle] nvarchar(500) NULL,
                        [VerticalAxisTitle] nvarchar(200) NOT NULL,
                        [YAxisUnit] nvarchar(50) NOT NULL,
                        [ShowLegend] bit NOT NULL
                            CONSTRAINT [DF_ChartTemplate_ShowLegend]
                            DEFAULT (1),
                        [LegendPosition] nvarchar(50) NOT NULL,
                        [PngWidthPixels] int NOT NULL,
                        [PngHeightPixels] int NOT NULL,
                        [DefaultStartDateMode] nvarchar(20) NOT NULL
                            CONSTRAINT [DF_ChartTemplate_DefaultStartDateMode]
                            DEFAULT (N'ReportStart'),
                        [WidthMm] int NOT NULL
                            CONSTRAINT [DF_ChartTemplate_WidthMm] DEFAULT (150),
                        [HeightMm] int NOT NULL
                            CONSTRAINT [DF_ChartTemplate_HeightMm] DEFAULT (80),
                        [ResolutionDpi] smallint NOT NULL
                            CONSTRAINT [DF_ChartTemplate_ResolutionDpi] DEFAULT (300),
                        [ChartFontFamily] nvarchar(100) NOT NULL
                            CONSTRAINT [DF_ChartTemplate_ChartFontFamily] DEFAULT (N'Arial'),
                        [DefaultLineWidth] decimal(10,3) NOT NULL,
                        [DefaultMarkerSize] decimal(10,3) NOT NULL,
                        [ShowGridLines] bit NOT NULL,
                        [IsDeleted] bit NOT NULL
                            CONSTRAINT [DF_ChartTemplate_IsDeleted]
                            DEFAULT (0),
                        [CreatedUtc] datetime2(3) NOT NULL
                            CONSTRAINT [DF_ChartTemplate_CreatedUtc]
                            DEFAULT (SYSUTCDATETIME()),
                        [UpdatedUtc] datetime2(3) NOT NULL
                            CONSTRAINT [DF_ChartTemplate_UpdatedUtc]
                            DEFAULT (SYSUTCDATETIME()),

                        CONSTRAINT [PK_ChartTemplate]
                            PRIMARY KEY CLUSTERED ([ChartTemplate_ID]),

                        CONSTRAINT [FK_ChartTemplate_Project]
                            FOREIGN KEY ([Project_ID])
                            REFERENCES [dbo].[Project] ([Project_ID])
                            ON DELETE NO ACTION
                            ON UPDATE NO ACTION,

                        CONSTRAINT [FK_ChartTemplate_ChartType]
                            FOREIGN KEY ([ChartType_ID])
                            REFERENCES [dbo].[ChartType] ([ChartType_ID])
                            ON DELETE NO ACTION
                            ON UPDATE NO ACTION,

                        CONSTRAINT [CK_ChartTemplate_Duration]
                            CHECK ([DefaultDurationDays] > 0),

                        CONSTRAINT [CK_ChartTemplate_Triggers]
                            CHECK
                            (
                                [GreenTrigger] > 0
                                AND [GreenTrigger] < [AmberTrigger]
                                AND [AmberTrigger] < [RedTrigger]
                            ),

                        CONSTRAINT [CK_ChartTemplate_VerticalAxis]
                            CHECK ([VerticalAxisMinimum] < [VerticalAxisMaximum])
                    );

                    CREATE UNIQUE INDEX
                        [UX_ChartTemplate_Active_Project_Name]
                        ON [dbo].[ChartTemplate]
                        (
                            [Project_ID],
                            [TemplateName]
                        )
                        WHERE [IsDeleted] = 0;
                END;

                IF COL_LENGTH(N'dbo.ChartTemplate', N'DefaultStartDateMode') IS NULL
                BEGIN
                    ALTER TABLE [dbo].[ChartTemplate]
                        ADD [DefaultStartDateMode] nvarchar(20) NOT NULL
                                CONSTRAINT [DF_ChartTemplate_DefaultStartDateMode]
                                DEFAULT (N'ReportStart'),
                            [WidthMm] int NOT NULL
                                CONSTRAINT [DF_ChartTemplate_WidthMm] DEFAULT (150),
                            [HeightMm] int NOT NULL
                                CONSTRAINT [DF_ChartTemplate_HeightMm] DEFAULT (80),
                            [ResolutionDpi] smallint NOT NULL
                                CONSTRAINT [DF_ChartTemplate_ResolutionDpi] DEFAULT (300),
                            [ChartFontFamily] nvarchar(100) NOT NULL
                                CONSTRAINT [DF_ChartTemplate_ChartFontFamily] DEFAULT (N'Arial');
                END;

                IF OBJECT_ID(N'dbo.ChartTemplateSeries', N'U') IS NULL
                BEGIN
                    CREATE TABLE [dbo].[ChartTemplateSeries]
                    (
                        [ChartTemplateSeries_ID] int IDENTITY(1,1) NOT NULL,
                        [ChartTemplate_ID] int NOT NULL,
                        [DataElementKey] nvarchar(100) NOT NULL,
                        [LegendText] nvarchar(300) NULL,
                        [LineWidth] decimal(10,3) NOT NULL,
                        [MarkerSize] decimal(10,3) NOT NULL,
                        [DisplayOrder] int NOT NULL,

                        CONSTRAINT [PK_ChartTemplateSeries]
                            PRIMARY KEY CLUSTERED ([ChartTemplateSeries_ID]),

                        CONSTRAINT [FK_ChartTemplateSeries_ChartTemplate]
                            FOREIGN KEY ([ChartTemplate_ID])
                            REFERENCES [dbo].[ChartTemplate] ([ChartTemplate_ID])
                            ON DELETE CASCADE
                            ON UPDATE NO ACTION,

                        CONSTRAINT [UQ_ChartTemplateSeries_Order]
                            UNIQUE ([ChartTemplate_ID], [DisplayOrder]),

                        CONSTRAINT [CK_ChartTemplateSeries_DisplayOrder]
                            CHECK ([DisplayOrder] > 0)
                    );
                END;
                """;

            #endregion


            #region Ensure Template Schema

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand command =
                new(
                    cmdText:
                        sql,
                    connection:
                        databaseConnection);

            await command.ExecuteNonQueryAsync();

            #endregion
        }


        private async Task RefreshChartTemplatesAsync(
            int? selectedTemplateId = null)
        {
            #region Validate Project

            _chartTemplates.Clear();

            if (!_activeProjectId.HasValue)
            {
                return;
            }

            #endregion


            #region Load Templates

            const string sql = """
                SELECT
                    T.[ChartTemplate_ID],
                    T.[TemplateName],
                    CT.[ChartTypeKey]
                FROM [dbo].[ChartTemplate] AS T
                INNER JOIN [dbo].[ChartType] AS CT
                    ON CT.[ChartType_ID] = T.[ChartType_ID]
                WHERE
                    T.[Project_ID] = @Project_ID
                    AND T.[IsDeleted] = 0
                ORDER BY
                    T.[TemplateName];
                """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand command =
                new(
                    cmdText:
                        sql,
                    connection:
                        databaseConnection);

            command.Parameters.Add(
                parameterName:
                    "@Project_ID",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    _activeProjectId.Value;

            await using SqlDataReader reader =
                await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                _chartTemplates.Add(
                    item:
                        new ChartTemplateUiItem
                        {
                            ChartTemplateId =
                                reader.GetInt32(0),

                            TemplateName =
                                reader.GetString(1),

                            ChartTypeKey =
                                reader.GetString(2)
                        });
            }

            #endregion


            #region Restore Selection

            if (selectedTemplateId.HasValue)
            {
                foreach (ChartTemplateUiItem item
                    in _chartTemplates)
                {
                    if (item.ChartTemplateId ==
                        selectedTemplateId.Value)
                    {
                        cmbChartTemplate.SelectedItem =
                            item;

                        break;
                    }
                }
            }

            #endregion
        }


        private async Task<ChartTemplateUiItem?> FindChartTemplateByNameAsync(
            string templateName)
        {
            if (!_activeProjectId.HasValue)
            {
                throw new InvalidOperationException(
                    "Select an active project.");
            }

            const string sql = """
                SELECT TOP (1)
                    T.[ChartTemplate_ID],
                    T.[TemplateName],
                    CT.[ChartTypeKey]
                FROM [dbo].[ChartTemplate] AS T
                INNER JOIN [dbo].[ChartType] AS CT
                    ON CT.[ChartType_ID] = T.[ChartType_ID]
                WHERE
                    T.[Project_ID] = @Project_ID
                    AND T.[IsDeleted] = 0
                    AND UPPER(T.[TemplateName]) = UPPER(@TemplateName);
                """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand command =
                new(
                    cmdText:
                        sql,
                    connection:
                        databaseConnection);

            command.Parameters.Add(
                parameterName:
                    "@Project_ID",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    _activeProjectId.Value;

            command.Parameters.Add(
                parameterName:
                    "@TemplateName",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    200)
                .Value =
                    templateName;

            await using SqlDataReader reader =
                await command.ExecuteReaderAsync();

            if (!await reader.ReadAsync())
            {
                return null;
            }

            return new ChartTemplateUiItem
            {
                ChartTemplateId =
                    reader.GetInt32(0),

                TemplateName =
                    reader.GetString(1),

                ChartTypeKey =
                    reader.GetString(2)
            };
        }


        private async Task<int> SaveChartTemplateAsync(
            int? chartTemplateId,
            string templateName)
        {
            #region Read Template Values

            if (!_activeProjectId.HasValue)
            {
                throw new InvalidOperationException(
                    "Select an active project.");
            }

            if (cmbChartType.SelectedItem
                is not ChartTypeUiItem chartType)
            {
                throw new InvalidOperationException(
                    "Select a chart type.");
            }

            int durationDays =
                int.TryParse(
                    s: txtChartRecentDays.Text,
                    style: NumberStyles.Integer,
                    provider: CultureInfo.InvariantCulture,
                    result: out int parsedDuration) &&
                parsedDuration > 0
                    ? parsedDuration
                    : DefaultChartRecentDays;

            decimal greenTrigger =
                decimal.Parse(
                    s: txtChartGreenTrigger.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            decimal amberTrigger =
                decimal.Parse(
                    s: txtChartAmberTrigger.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            decimal redTrigger =
                decimal.Parse(
                    s: txtChartRedTrigger.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            if (chkChartAutomaticYAxis.IsChecked == true)
            {
                ApplyDefaultVerticalAxisFromRedTrigger();
            }

            decimal verticalMinimum =
                decimal.Parse(
                    s: txtChartYAxisMinimum.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            decimal verticalMaximum =
                decimal.Parse(
                    s: txtChartYAxisMaximum.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            int pngWidth =
                int.Parse(
                    s: txtChartPngWidth.Text,
                    provider: CultureInfo.InvariantCulture);

            int pngHeight =
                int.Parse(
                    s: txtChartPngHeight.Text,
                    provider: CultureInfo.InvariantCulture);

            double defaultLineWidth =
                ParsePositiveDoubleOrDefault(
                    text:
                        txtChartDefaultLineWidth.Text,
                    defaultValue:
                        2.0);

            double defaultMarkerSize =
                ParseNonNegativeDoubleOrDefault(
                    text:
                        txtChartDefaultMarkerSize.Text,
                    defaultValue:
                        4.0);

            #endregion


            #region Open Transaction

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            using SqlTransaction transaction =
                databaseConnection.BeginTransaction();

            try
            {
                #region Resolve Chart Type

                const string chartTypeSql = """
                    SELECT [ChartType_ID]
                    FROM [dbo].[ChartType]
                    WHERE [ChartTypeKey] = @ChartTypeKey;
                    """;

                await using SqlCommand chartTypeCommand =
                    new(
                        cmdText:
                            chartTypeSql,
                        connection:
                            databaseConnection,
                        transaction:
                            transaction);

                chartTypeCommand.Parameters.Add(
                    parameterName:
                        "@ChartTypeKey",
                    sqlDbType:
                        System.Data.SqlDbType.NVarChar,
                    size:
                        100)
                    .Value =
                        chartType.Key;

                object? chartTypeIdValue =
                    await chartTypeCommand.ExecuteScalarAsync();

                if (chartTypeIdValue is null ||
                    chartTypeIdValue == DBNull.Value)
                {
                    throw new InvalidOperationException(
                        "Chart type is not registered in the database.");
                }

                int chartTypeId =
                    Convert.ToInt32(
                        value:
                            chartTypeIdValue,
                        provider:
                            CultureInfo.InvariantCulture);

                #endregion


                #region Insert Or Update Template

                int resolvedTemplateId;

                if (chartTemplateId.HasValue)
                {
                    const string updateSql = """
                        UPDATE [dbo].[ChartTemplate]
                        SET
                            [TemplateName] = @TemplateName,
                            [ChartType_ID] = @ChartType_ID,
                            [DefaultDurationDays] = @DefaultDurationDays,
                            [GreenTrigger] = @GreenTrigger,
                            [AmberTrigger] = @AmberTrigger,
                            [RedTrigger] = @RedTrigger,
                            [UseAutomaticVerticalAxis] = @UseAutomaticVerticalAxis,
                            [VerticalAxisMinimum] = @VerticalAxisMinimum,
                            [VerticalAxisMaximum] = @VerticalAxisMaximum,
                            [ChartTitle] = @ChartTitle,
                            [VerticalAxisTitle] = @VerticalAxisTitle,
                            [YAxisUnit] = @YAxisUnit,
                            [ShowLegend] = @ShowLegend,
                            [LegendPosition] = @LegendPosition,
                            [PngWidthPixels] = @PngWidthPixels,
                            [PngHeightPixels] = @PngHeightPixels,
                            [WidthMm] = @WidthMm,
                            [HeightMm] = @HeightMm,
                            [ResolutionDpi] = @ResolutionDpi,
                            [ChartFontFamily] = N'Arial',
                            [DefaultStartDateMode] = @StartDateMode,
                            [DefaultLineWidth] = @DefaultLineWidth,
                            [DefaultMarkerSize] = @DefaultMarkerSize,
                            [ShowGridLines] = @ShowGridLines,
                            [UpdatedUtc] = SYSUTCDATETIME()
                        WHERE
                            [ChartTemplate_ID] = @ChartTemplate_ID
                            AND [Project_ID] = @Project_ID
                            AND [IsDeleted] = 0;
                        """;

                    await using SqlCommand updateCommand =
                        new(
                            cmdText:
                                updateSql,
                            connection:
                                databaseConnection,
                            transaction:
                                transaction);

                    AddChartTemplateParameters(
                        command:
                            updateCommand,
                        templateName:
                            templateName,
                        chartTypeId:
                            chartTypeId,
                        durationDays:
                            durationDays,
                        greenTrigger:
                            greenTrigger,
                        amberTrigger:
                            amberTrigger,
                        redTrigger:
                            redTrigger,
                        verticalMinimum:
                            verticalMinimum,
                        verticalMaximum:
                            verticalMaximum,
                        pngWidth:
                            pngWidth,
                        pngHeight:
                            pngHeight,
                        defaultLineWidth:
                            defaultLineWidth,
                        defaultMarkerSize:
                            defaultMarkerSize);

                    updateCommand.Parameters.Add(
                        parameterName:
                            "@ChartTemplate_ID",
                        sqlDbType:
                            System.Data.SqlDbType.Int)
                        .Value =
                            chartTemplateId.Value;

                    int rowsUpdated =
                        await updateCommand.ExecuteNonQueryAsync();

                    if (rowsUpdated != 1)
                    {
                        throw new InvalidOperationException(
                            "Template update failed.");
                    }

                    resolvedTemplateId =
                        chartTemplateId.Value;
                }
                else
                {
                    const string insertSql = """
                        INSERT INTO [dbo].[ChartTemplate]
                        (
                            [Project_ID],
                            [TemplateName],
                            [ChartType_ID],
                            [DefaultDurationDays],
                            [GreenTrigger],
                            [AmberTrigger],
                            [RedTrigger],
                            [UseAutomaticVerticalAxis],
                            [VerticalAxisMinimum],
                            [VerticalAxisMaximum],
                            [ChartTitle],
                            [VerticalAxisTitle],
                            [YAxisUnit],
                            [ShowLegend],
                            [LegendPosition],
                            [PngWidthPixels],
                            [PngHeightPixels],
                            [WidthMm],
                            [HeightMm],
                            [ResolutionDpi],
                            [ChartFontFamily],
                            [DefaultStartDateMode],
                            [DefaultLineWidth],
                            [DefaultMarkerSize],
                            [ShowGridLines],
                            [IsDeleted]
                        )
                        OUTPUT INSERTED.[ChartTemplate_ID]
                        VALUES
                        (
                            @Project_ID,
                            @TemplateName,
                            @ChartType_ID,
                            @DefaultDurationDays,
                            @GreenTrigger,
                            @AmberTrigger,
                            @RedTrigger,
                            @UseAutomaticVerticalAxis,
                            @VerticalAxisMinimum,
                            @VerticalAxisMaximum,
                            @ChartTitle,
                            @VerticalAxisTitle,
                            @YAxisUnit,
                            @ShowLegend,
                            @LegendPosition,
                            @PngWidthPixels,
                            @PngHeightPixels,
                            @WidthMm,
                            @HeightMm,
                            @ResolutionDpi,
                            N'Arial',
                            @StartDateMode,
                            @DefaultLineWidth,
                            @DefaultMarkerSize,
                            @ShowGridLines,
                            0
                        );
                        """;

                    await using SqlCommand insertCommand =
                        new(
                            cmdText:
                                insertSql,
                            connection:
                                databaseConnection,
                            transaction:
                                transaction);

                    AddChartTemplateParameters(
                        command:
                            insertCommand,
                        templateName:
                            templateName,
                        chartTypeId:
                            chartTypeId,
                        durationDays:
                            durationDays,
                        greenTrigger:
                            greenTrigger,
                        amberTrigger:
                            amberTrigger,
                        redTrigger:
                            redTrigger,
                        verticalMinimum:
                            verticalMinimum,
                        verticalMaximum:
                            verticalMaximum,
                        pngWidth:
                            pngWidth,
                        pngHeight:
                            pngHeight,
                        defaultLineWidth:
                            defaultLineWidth,
                        defaultMarkerSize:
                            defaultMarkerSize);

                    object? newTemplateIdValue =
                        await insertCommand.ExecuteScalarAsync();

                    if (newTemplateIdValue is null ||
                        newTemplateIdValue == DBNull.Value)
                    {
                        throw new InvalidOperationException(
                            "Template insert failed.");
                    }

                    resolvedTemplateId =
                        Convert.ToInt32(
                            value:
                                newTemplateIdValue,
                        provider:
                            CultureInfo.InvariantCulture);
                }

                #endregion


                #region Replace Template Series

                const string deleteSeriesSql = """
                    DELETE FROM [dbo].[ChartTemplateSeries]
                    WHERE [ChartTemplate_ID] = @ChartTemplate_ID;
                    """;

                await using (SqlCommand deleteSeriesCommand =
                    new(
                        cmdText:
                            deleteSeriesSql,
                        connection:
                            databaseConnection,
                        transaction:
                            transaction))
                {
                    deleteSeriesCommand.Parameters.Add(
                        parameterName:
                            "@ChartTemplate_ID",
                        sqlDbType:
                            System.Data.SqlDbType.Int)
                        .Value =
                            resolvedTemplateId;

                    await deleteSeriesCommand.ExecuteNonQueryAsync();
                }

                const string insertSeriesSql = """
                    INSERT INTO [dbo].[ChartTemplateSeries]
                    (
                        [ChartTemplate_ID],
                        [DataElementKey],
                        [LegendText],
                        [LineWidth],
                        [MarkerSize],
                        [DisplayOrder]
                    )
                    VALUES
                    (
                        @ChartTemplate_ID,
                        @DataElementKey,
                        @LegendText,
                        @LineWidth,
                        @MarkerSize,
                        @DisplayOrder
                    );
                    """;

                foreach (ChartSeriesUiItem series
                    in _chartSeries.OrderBy(
                        keySelector:
                            item => item.DisplayOrder))
                {
                    await using SqlCommand seriesCommand =
                        new(
                            cmdText:
                                insertSeriesSql,
                            connection:
                                databaseConnection,
                            transaction:
                                transaction);

                    seriesCommand.Parameters.Add(
                        parameterName:
                            "@ChartTemplate_ID",
                        sqlDbType:
                            System.Data.SqlDbType.Int)
                        .Value =
                            resolvedTemplateId;

                    seriesCommand.Parameters.Add(
                        parameterName:
                            "@DataElementKey",
                        sqlDbType:
                            System.Data.SqlDbType.NVarChar,
                        size:
                            100)
                        .Value =
                            series.DataElementDisplayName;

                    seriesCommand.Parameters.Add(
                        parameterName:
                            "@LegendText",
                        sqlDbType:
                            System.Data.SqlDbType.NVarChar,
                        size:
                            300)
                        .Value =
                            string.IsNullOrWhiteSpace(
                                value:
                                    series.LegendText)
                                ? DBNull.Value
                                : series.LegendText;

                    SqlParameter lineWidthParameter =
                        seriesCommand.Parameters.Add(
                            parameterName:
                                "@LineWidth",
                            sqlDbType:
                                System.Data.SqlDbType.Decimal);

                    lineWidthParameter.Precision =
                        10;

                    lineWidthParameter.Scale =
                        3;

                    lineWidthParameter.Value =
                        Convert.ToDecimal(
                            value:
                                series.LineWidth,
                            provider:
                                CultureInfo.InvariantCulture);

                    SqlParameter markerSizeParameter =
                        seriesCommand.Parameters.Add(
                            parameterName:
                                "@MarkerSize",
                            sqlDbType:
                                System.Data.SqlDbType.Decimal);

                    markerSizeParameter.Precision =
                        10;

                    markerSizeParameter.Scale =
                        3;

                    markerSizeParameter.Value =
                        Convert.ToDecimal(
                            value:
                                series.MarkerSize,
                            provider:
                                CultureInfo.InvariantCulture);

                    seriesCommand.Parameters.Add(
                        parameterName:
                            "@DisplayOrder",
                        sqlDbType:
                            System.Data.SqlDbType.Int)
                        .Value =
                            series.DisplayOrder;

                    await seriesCommand.ExecuteNonQueryAsync();
                }

                #endregion


                transaction.Commit();

                return resolvedTemplateId;
            }
            catch
            {
                try
                {
                    transaction.Rollback();
                }
                catch
                {
                    // Preserve original exception.
                }

                throw;
            }

            #endregion
        }


        private void AddChartTemplateParameters(
            SqlCommand command,
            string templateName,
            int chartTypeId,
            int durationDays,
            decimal greenTrigger,
            decimal amberTrigger,
            decimal redTrigger,
            decimal verticalMinimum,
            decimal verticalMaximum,
            int pngWidth,
            int pngHeight,
            double defaultLineWidth,
            double defaultMarkerSize)
        {
            command.Parameters.Add(
                parameterName:
                    "@Project_ID",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    _activeProjectId!.Value;

            command.Parameters.Add(
                parameterName:
                    "@TemplateName",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    200)
                .Value =
                    templateName;

            command.Parameters.Add(
                parameterName:
                    "@ChartType_ID",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    chartTypeId;

            command.Parameters.Add(
                parameterName:
                    "@DefaultDurationDays",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    durationDays;

            AddDecimalParameter(
                command:
                    command,
                name:
                    "@GreenTrigger",
                value:
                    greenTrigger);

            AddDecimalParameter(
                command:
                    command,
                name:
                    "@AmberTrigger",
                value:
                    amberTrigger);

            AddDecimalParameter(
                command:
                    command,
                name:
                    "@RedTrigger",
                value:
                    redTrigger);

            command.Parameters.Add(
                parameterName:
                    "@UseAutomaticVerticalAxis",
                sqlDbType:
                    System.Data.SqlDbType.Bit)
                .Value =
                    chkChartAutomaticYAxis.IsChecked == true;

            AddDecimalParameter(
                command:
                    command,
                name:
                    "@VerticalAxisMinimum",
                value:
                    verticalMinimum);

            AddDecimalParameter(
                command:
                    command,
                name:
                    "@VerticalAxisMaximum",
                value:
                    verticalMaximum);

            command.Parameters.Add(
                parameterName:
                    "@ChartTitle",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    500)
                .Value =
                    string.IsNullOrWhiteSpace(
                        value:
                            txtChartTitleOverride.Text)
                        ? DBNull.Value
                        : txtChartTitleOverride.Text.Trim();

            command.Parameters.Add(
                parameterName:
                    "@VerticalAxisTitle",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    200)
                .Value =
                    txtChartYAxisTitle.Text.Trim();

            command.Parameters.Add(
                parameterName:
                    "@YAxisUnit",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    50)
                .Value =
                    txtChartUnit.Text.Trim();

            command.Parameters.Add(
                parameterName:
                    "@ShowLegend",
                sqlDbType:
                    System.Data.SqlDbType.Bit)
                .Value =
                    chkChartShowLegend.IsChecked == true;

            command.Parameters.Add(
                parameterName:
                    "@LegendPosition",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    50)
                .Value =
                    GetSelectedLegendPosition();

            command.Parameters.Add(
                parameterName:
                    "@PngWidthPixels",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    pngWidth;

            command.Parameters.Add(
                parameterName:
                    "@PngHeightPixels",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    pngHeight;

            command.Parameters.Add(
                parameterName: "@WidthMm",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value = int.Parse(
                    s: txtChartWidthMm.Text,
                    provider: CultureInfo.InvariantCulture);

            command.Parameters.Add(
                parameterName: "@HeightMm",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value = int.Parse(
                    s: txtChartHeightMm.Text,
                    provider: CultureInfo.InvariantCulture);

            command.Parameters.Add(
                parameterName: "@ResolutionDpi",
                sqlDbType: System.Data.SqlDbType.SmallInt)
                .Value = GetSelectedChartResolutionDpi();

            command.Parameters.Add(
                parameterName: "@StartDateMode",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 20)
                .Value = GetSelectedChartStartDateMode();

            SqlParameter lineWidthParameter =
                command.Parameters.Add(
                    parameterName:
                        "@DefaultLineWidth",
                    sqlDbType:
                        System.Data.SqlDbType.Decimal);

            lineWidthParameter.Precision =
                10;

            lineWidthParameter.Scale =
                3;

            lineWidthParameter.Value =
                Convert.ToDecimal(
                    value:
                        defaultLineWidth,
                    provider:
                        CultureInfo.InvariantCulture);

            SqlParameter markerSizeParameter =
                command.Parameters.Add(
                    parameterName:
                        "@DefaultMarkerSize",
                    sqlDbType:
                        System.Data.SqlDbType.Decimal);

            markerSizeParameter.Precision =
                10;

            markerSizeParameter.Scale =
                3;

            markerSizeParameter.Value =
                Convert.ToDecimal(
                    value:
                        defaultMarkerSize,
                    provider:
                        CultureInfo.InvariantCulture);

            command.Parameters.Add(
                parameterName:
                    "@ShowGridLines",
                sqlDbType:
                    System.Data.SqlDbType.Bit)
                .Value =
                    chkChartGridLines.IsChecked == true;
        }


        private static void AddDecimalParameter(
            SqlCommand command,
            string name,
            decimal value)
        {
            SqlParameter parameter =
                command.Parameters.Add(
                    parameterName:
                        name,
                    sqlDbType:
                        System.Data.SqlDbType.Decimal);

            parameter.Precision =
                18;

            parameter.Scale =
                6;

            parameter.Value =
                value;
        }


        private async Task LoadChartTemplateIntoEditorAsync(
            int chartTemplateId)
        {
            #region Load Template Header

            const string sql = """
                SELECT
                    T.[ChartTemplate_ID],
                    T.[TemplateName],
                    CT.[ChartTypeKey],
                    T.[DefaultDurationDays],
                    T.[GreenTrigger],
                    T.[AmberTrigger],
                    T.[RedTrigger],
                    T.[UseAutomaticVerticalAxis],
                    T.[VerticalAxisMinimum],
                    T.[VerticalAxisMaximum],
                    T.[ChartTitle],
                    T.[VerticalAxisTitle],
                    T.[ShowLegend],
                    T.[LegendPosition],
                    T.[PngWidthPixels],
                    T.[PngHeightPixels],
                    T.[DefaultLineWidth],
                    T.[DefaultMarkerSize],
                    T.[ShowGridLines],
                    T.[WidthMm],
                    T.[HeightMm],
                    T.[ResolutionDpi],
                    T.[DefaultStartDateMode]
                FROM [dbo].[ChartTemplate] AS T
                INNER JOIN [dbo].[ChartType] AS CT
                    ON CT.[ChartType_ID] = T.[ChartType_ID]
                WHERE
                    T.[ChartTemplate_ID] = @ChartTemplate_ID
                    AND T.[Project_ID] = @Project_ID
                    AND T.[IsDeleted] = 0;
                """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand command =
                new(
                    cmdText:
                        sql,
                    connection:
                        databaseConnection);

            command.Parameters.Add(
                parameterName:
                    "@ChartTemplate_ID",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    chartTemplateId;

            command.Parameters.Add(
                parameterName:
                    "@Project_ID",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    _activeProjectId
                    ?? throw new InvalidOperationException(
                        "Select an active project.");

            await using SqlDataReader reader =
                await command.ExecuteReaderAsync();

            if (!await reader.ReadAsync())
            {
                throw new InvalidOperationException(
                    "Template not found.");
            }

            string templateName =
                reader.GetString(1);

            string chartTypeKey =
                reader.GetString(2);

            int durationDays =
                reader.GetInt32(3);

            decimal greenTrigger =
                reader.GetDecimal(4);

            decimal amberTrigger =
                reader.GetDecimal(5);

            decimal redTrigger =
                reader.GetDecimal(6);

            bool automaticVerticalAxis =
                reader.GetBoolean(7);

            decimal verticalMinimum =
                reader.GetDecimal(8);

            decimal verticalMaximum =
                reader.GetDecimal(9);

            string chartTitle =
                reader.IsDBNull(10)
                    ? string.Empty
                    : reader.GetString(10);

            string verticalAxisTitle =
                reader.GetString(11);

            bool showLegend =
                reader.GetBoolean(12);

            string legendPosition =
                reader.GetString(13);

            int pngWidth =
                reader.GetInt32(14);

            int pngHeight =
                reader.GetInt32(15);

            decimal defaultLineWidth =
                reader.GetDecimal(16);

            decimal defaultMarkerSize =
                reader.GetDecimal(17);

            bool showGridLines =
                reader.GetBoolean(18);

            int widthMm =
                reader.GetInt32(19);

            int heightMm =
                reader.GetInt32(20);

            int resolutionDpi =
                reader.GetInt16(21);

            string startDateMode =
                reader.GetString(22);

            await reader.DisposeAsync();

            #endregion


            #region Apply Template To New Chart

            _loadedChartTemplateId =
                chartTemplateId;

            _loadedChartDefinitionId =
                null;

            _loadedChartNumber =
                null;

            cmbExistingChart.SelectedItem =
                null;

            txtChartNumber.Text =
                "New";

            txtChartName.Text =
                templateName;

            foreach (object item
                in cmbChartType.Items)
            {
                if (item is ChartTypeUiItem chartType &&
                    string.Equals(
                        a:
                            chartType.Key,
                        b:
                            chartTypeKey,
                        comparisonType:
                            StringComparison.OrdinalIgnoreCase))
                {
                    cmbChartType.SelectedItem =
                        chartType;

                    break;
                }
            }

            txtChartRecentDays.Text =
                durationDays.ToString(
                    provider:
                        CultureInfo.InvariantCulture);

            DateTime today =
                DateTime.Today;

            dpChartEndDate.SelectedDate =
                today;

            dpChartStartDate.SelectedDate =
                today.AddDays(
                    value:
                        -durationDays);

            txtChartGreenTrigger.Text =
                greenTrigger.ToString(
                    provider:
                        CultureInfo.InvariantCulture);

            txtChartAmberTrigger.Text =
                amberTrigger.ToString(
                    provider:
                        CultureInfo.InvariantCulture);

            txtChartRedTrigger.Text =
                redTrigger.ToString(
                    provider:
                        CultureInfo.InvariantCulture);

            chkChartAutomaticYAxis.IsChecked =
                automaticVerticalAxis;

            txtChartYAxisMinimum.Text =
                verticalMinimum.ToString(
                    provider:
                        CultureInfo.InvariantCulture);

            txtChartYAxisMaximum.Text =
                verticalMaximum.ToString(
                    provider:
                        CultureInfo.InvariantCulture);

            txtChartTitleOverride.Text =
                chartTitle;

            txtChartYAxisTitle.Text =
                verticalAxisTitle;

            chkChartShowLegend.IsChecked =
                showLegend;

            SelectLegendPosition(
                legendPosition:
                    legendPosition);

            txtChartPngWidth.Text =
                pngWidth.ToString(
                    provider:
                        CultureInfo.InvariantCulture);

            txtChartPngHeight.Text =
                pngHeight.ToString(
                    provider:
                        CultureInfo.InvariantCulture);

            txtChartWidthMm.Text =
                widthMm.ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartHeightMm.Text =
                heightMm.ToString(
                    provider: CultureInfo.InvariantCulture);

            SelectChartResolutionDpi(
                resolutionDpi: resolutionDpi);

            SelectChartStartDateMode(
                startDateMode: startDateMode);

            UpdateChartPixelDimensions();

            txtChartDefaultLineWidth.Text =
                defaultLineWidth.ToString(
                    provider:
                        CultureInfo.InvariantCulture);

            txtChartDefaultMarkerSize.Text =
                defaultMarkerSize.ToString(
                    provider:
                        CultureInfo.InvariantCulture);

            chkChartGridLines.IsChecked =
                showGridLines;

            #endregion


            #region Load Template Series Definitions

            _chartSeries.Clear();

            const string seriesSql = """
                SELECT
                    [DataElementKey],
                    [LegendText],
                    [LineWidth],
                    [MarkerSize],
                    [DisplayOrder]
                FROM [dbo].[ChartTemplateSeries]
                WHERE [ChartTemplate_ID] = @ChartTemplate_ID
                ORDER BY [DisplayOrder];
                """;

            await using SqlCommand seriesCommand =
                new(
                    cmdText:
                        seriesSql,
                    connection:
                        databaseConnection);

            seriesCommand.Parameters.Add(
                parameterName:
                    "@ChartTemplate_ID",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    chartTemplateId;

            await using SqlDataReader seriesReader =
                await seriesCommand.ExecuteReaderAsync();

            while (await seriesReader.ReadAsync())
            {
                string dataElement =
                    seriesReader.GetString(0);

                _chartSeries.Add(
                    item:
                        new ChartSeriesUiItem
                        {
                            DisplayOrder =
                                seriesReader.GetInt32(4),

                            EntityDisplayName =
                                "<Select entity>",

                            DataElementDisplayName =
                                dataElement,

                            LegendText =
                                seriesReader.IsDBNull(1)
                                    ? dataElement
                                    : seriesReader.GetString(1),

                            LineWidth =
                                Convert.ToDouble(
                                    value:
                                        seriesReader.GetDecimal(2),
                                    provider:
                                        CultureInfo.InvariantCulture),

                            MarkerSize =
                                Convert.ToDouble(
                                    value:
                                        seriesReader.GetDecimal(3),
                                    provider:
                                        CultureInfo.InvariantCulture)
                        });
            }

            #endregion
        }

        #endregion


        #region Chart Repository Persistence

        private async Task EnsureChartTypeCatalogueAsync()
        {
            #region Open Database

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            #endregion


            #region Ensure Each Chart Type Exists

            const string mergeSql = """
                IF NOT EXISTS
                (
                    SELECT 1
                    FROM [dbo].[ChartType]
                    WHERE [ChartTypeKey] = @ChartTypeKey
                )
                BEGIN
                    INSERT INTO [dbo].[ChartType]
                    (
                        [ChartTypeKey],
                        [DisplayName],
                        [DataSourceTable],
                        [TimestampColumnName],
                        [EntityIdColumnName],
                        [EntityKind],
                        [DefaultTitleTemplate],
                        [DefaultYAxisTitle],
                        [DefaultXAxisTitle],
                        [SupportsReferenceAdjustment],
                        [SupportsTriggerBands],
                        [DefaultTriggerBandsSymmetric],
                        [IsEnabled],
                        [DisplayOrder]
                    )
                    VALUES
                    (
                        @ChartTypeKey,
                        @DisplayName,
                        @DataSourceTable,
                        N'UTCtime',
                        @EntityIdColumnName,
                        @EntityKind,
                        N'{ChartType} - {Sensors}',
                        @DefaultYAxisTitle,
                        N'Date / Time',
                        0,
                        1,
                        1,
                        1,
                        @DisplayOrder
                    );
                END
                ELSE
                BEGIN
                    UPDATE [dbo].[ChartType]
                    SET
                        [DisplayName] = @DisplayName,
                        [DataSourceTable] = @DataSourceTable,
                        [TimestampColumnName] = N'UTCtime',
                        [EntityIdColumnName] = @EntityIdColumnName,
                        [EntityKind] = @EntityKind,
                        [DefaultYAxisTitle] = @DefaultYAxisTitle,
                        [SupportsReferenceAdjustment] = 0,
                        [SupportsTriggerBands] = 1,
                        [DefaultTriggerBandsSymmetric] = 1,
                        [IsEnabled] = 1,
                        [DisplayOrder] = @DisplayOrder
                    WHERE [ChartTypeKey] = @ChartTypeKey;
                END;
                """;

            int displayOrder =
                0;

            foreach (ChartTypeUiItem chartType
                in _chartTypes.OrderBy(
                    keySelector: item => item.DisplayName,
                    comparer: StringComparer.OrdinalIgnoreCase))
            {
                displayOrder++;

                await using SqlCommand command =
                    new(
                        cmdText: mergeSql,
                        connection: databaseConnection);

                command.Parameters.Add(
                    parameterName: "@ChartTypeKey",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 100)
                    .Value =
                        chartType.Key;

                command.Parameters.Add(
                    parameterName: "@DisplayName",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 200)
                    .Value =
                        chartType.DisplayName;

                command.Parameters.Add(
                    parameterName: "@DataSourceTable",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 128)
                    .Value =
                        chartType.SourceTable;

                command.Parameters.Add(
                    parameterName: "@EntityIdColumnName",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 128)
                    .Value =
                        GetChartEntityIdColumnName(
                            entityKind:
                                chartType.EntityKind);

                command.Parameters.Add(
                    parameterName: "@EntityKind",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 30)
                    .Value =
                        chartType.EntityKind;

                command.Parameters.Add(
                    parameterName: "@DefaultYAxisTitle",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 200)
                    .Value =
                        chartType.VerticalAxisTitle;

                command.Parameters.Add(
                    parameterName: "@DisplayOrder",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value =
                        displayOrder;

                await command.ExecuteNonQueryAsync();
            }

            #endregion
        }


        private static string GetChartEntityIdColumnName(
            string entityKind)
        {
            return entityKind switch
            {
                "Sensor" => "SensorID",
                "Point" => "PointName_ID",
                "PrismPair" => "PrismPair_ID",
                "Track" => "Track_ID",
                "PrismArray" => "Array_ID",

                _ => throw new InvalidOperationException(
                    $"Unsupported chart entity kind '{entityKind}'.")
            };
        }


        #region Displacement Chart Engineering Conversion

        private static double CalculateDisplacementChartValueMm(
            string dataElement,
            decimal currentE,
            decimal currentN,
            decimal referenceE,
            decimal referenceN)
        {
            #region Calculate Coordinate Differences

            double deltaE =
                (double)(currentE - referenceE);

            double deltaN =
                (double)(currentN - referenceN);

            #endregion


            #region Convert Selected Element To Millimetres

            return dataElement switch
            {
                "dE" =>
                    deltaE * 1000.0,

                "dN" =>
                    deltaN * 1000.0,

                "d2D" =>
                    Math.Sqrt(
                        (deltaE * deltaE) +
                        (deltaN * deltaN)) *
                    1000.0,

                _ => throw new InvalidOperationException(
                    $"Unsupported Displacement data element '{dataElement}'.")
            };

            #endregion
        }

        #endregion


        private async Task RefreshExistingChartsAsync(
            int? selectedChartDefinitionId = null)
        {
            #region Clear Existing Chart List

            _existingCharts.Clear();

            if (!_activeProjectId.HasValue)
            {
                return;
            }

            #endregion


            #region Load Existing Charts

            const string sql = """
                SELECT
                    CD.[ChartDefinition_ID],
                    CD.[ChartNumber],
                    CD.[ChartName],
                    CT.[ChartTypeKey]
                FROM [dbo].[ChartDefinition] AS CD
                INNER JOIN [dbo].[ChartType] AS CT
                    ON CT.[ChartType_ID] = CD.[ChartType_ID]
                WHERE
                    CD.[Project_ID] = @Project_ID
                    AND CD.[IsDeleted] = 0
                ORDER BY
                    CD.[ChartNumber],
                    CD.[ChartName];
                """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand command =
                new(
                    cmdText: sql,
                    connection: databaseConnection);

            command.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    _activeProjectId.Value;

            await using SqlDataReader reader =
                await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                ExistingChartUiItem item =
                    new()
                    {
                        ChartDefinitionId =
                            reader.GetInt32(0),

                        ChartNumber =
                            reader.GetInt32(1),

                        ChartName =
                            reader.GetString(2),

                        ChartTypeKey =
                            reader.GetString(3)
                    };

                _existingCharts.Add(
                    item: item);
            }

            #endregion


            #region Restore Selection

            if (selectedChartDefinitionId.HasValue)
            {
                foreach (ExistingChartUiItem item
                    in _existingCharts)
                {
                    if (item.ChartDefinitionId ==
                        selectedChartDefinitionId.Value)
                    {
                        cmbExistingChart.SelectedItem =
                            item;

                        break;
                    }
                }
            }

            #endregion
        }


        private async Task<ExistingChartUiItem?> FindChartByNameAsync(
            string chartName,
            int? excludeChartDefinitionId)
        {
            #region Validate Lookup

            if (!_activeProjectId.HasValue)
            {
                throw new InvalidOperationException(
                    "Select an active project.");
            }

            string validatedChartName =
                chartName.Trim();

            #endregion


            #region Find Duplicate

            const string sql = """
                SELECT TOP (1)
                    CD.[ChartDefinition_ID],
                    CD.[ChartNumber],
                    CD.[ChartName],
                    CT.[ChartTypeKey]
                FROM [dbo].[ChartDefinition] AS CD
                INNER JOIN [dbo].[ChartType] AS CT
                    ON CT.[ChartType_ID] = CD.[ChartType_ID]
                WHERE
                    CD.[Project_ID] = @Project_ID
                    AND CD.[IsDeleted] = 0
                    AND UPPER(CD.[ChartName]) = UPPER(@ChartName)
                    AND
                    (
                        @ExcludeChartDefinition_ID IS NULL
                        OR CD.[ChartDefinition_ID] <> @ExcludeChartDefinition_ID
                    );
                """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand command =
                new(
                    cmdText: sql,
                    connection: databaseConnection);

            command.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    _activeProjectId.Value;

            command.Parameters.Add(
                parameterName: "@ChartName",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 200)
                .Value =
                    validatedChartName;

            command.Parameters.Add(
                parameterName: "@ExcludeChartDefinition_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    excludeChartDefinitionId.HasValue
                        ? excludeChartDefinitionId.Value
                        : DBNull.Value;

            await using SqlDataReader reader =
                await command.ExecuteReaderAsync();

            if (!await reader.ReadAsync())
            {
                return null;
            }

            return new ExistingChartUiItem
            {
                ChartDefinitionId =
                    reader.GetInt32(0),

                ChartNumber =
                    reader.GetInt32(1),

                ChartName =
                    reader.GetString(2),

                ChartTypeKey =
                    reader.GetString(3)
            };

            #endregion
        }


        private async Task LoadChartDefinitionIntoEditorAsync(
            int chartDefinitionId)
        {
            #region Define Chart Query

            const string chartSql = """
                SELECT
                    CD.[ChartDefinition_ID],
                    CD.[ChartNumber],
                    CD.[ChartName],
                    CT.[ChartTypeKey],
                    CD.[YAxisTitle],
                    CD.[YAxisUnit],
                    CD.[UseAutomaticYAxis],
                    CD.[FixedYAxisMinimum],
                    CD.[FixedYAxisMaximum],
                    CD.[ShowLegend],
                    CD.[LegendPosition],
                    CD.[PngWidthPixels],
                    CD.[PngHeightPixels],
                    CD.[WidthMm],
                    CD.[HeightMm],
                    CD.[ResolutionDpi],
                    CD.[StartDateMode],
                    CD.[AbsoluteStartUtc],
                    CD.[AbsoluteEndUtc],
                    CD.[TitleOverride]
                FROM [dbo].[ChartDefinition] AS CD
                INNER JOIN [dbo].[ChartType] AS CT
                    ON CT.[ChartType_ID] = CD.[ChartType_ID]
                WHERE
                    CD.[ChartDefinition_ID] = @ChartDefinition_ID
                    AND CD.[Project_ID] = @Project_ID
                    AND CD.[IsDeleted] = 0;
                """;

            #endregion


            #region Load Chart Header

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand chartCommand =
                new(
                    cmdText: chartSql,
                    connection: databaseConnection);

            chartCommand.Parameters.Add(
                parameterName: "@ChartDefinition_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    chartDefinitionId;

            chartCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    _activeProjectId
                    ?? throw new InvalidOperationException(
                        "Select an active project.");

            await using SqlDataReader reader =
                await chartCommand.ExecuteReaderAsync();

            if (!await reader.ReadAsync())
            {
                throw new InvalidOperationException(
                    "Chart definition not found.");
            }

            int chartNumber =
                reader.GetInt32(1);

            string chartName =
                reader.GetString(2);

            string chartTypeKey =
                reader.GetString(3);

            string verticalAxisTitle =
                reader.GetString(4);

            bool automaticVerticalAxis =
                reader.GetBoolean(6);

            decimal? fixedMinimum =
                reader.IsDBNull(7)
                    ? null
                    : reader.GetDecimal(7);

            decimal? fixedMaximum =
                reader.IsDBNull(8)
                    ? null
                    : reader.GetDecimal(8);

            bool showLegend =
                reader.GetBoolean(9);

            string legendPosition =
                reader.GetString(10);

            int pngWidth =
                reader.GetInt32(11);

            int pngHeight =
                reader.GetInt32(12);

            int widthMm =
                reader.GetInt32(13);

            int heightMm =
                reader.GetInt32(14);

            int resolutionDpi =
                reader.GetInt16(15);

            string startDateMode =
                reader.GetString(16);

            DateTime absoluteStart =
                reader.GetDateTime(17);

            DateTime absoluteEndExclusive =
                reader.GetDateTime(18);

            string chartTitle =
                reader.IsDBNull(19)
                    ? string.Empty
                    : reader.GetString(19);

            await reader.DisposeAsync();

            #endregion


            #region Apply Chart Header To Editor

            _loadedChartDefinitionId =
                chartDefinitionId;

            _loadedChartNumber =
                chartNumber;

            txtChartNumber.Text =
                $"Chart_{chartNumber:0000}";

            txtChartName.Text =
                chartName;

            foreach (ChartTypeUiItem chartType
                in _chartTypes)
            {
                if (string.Equals(
                    a: chartType.Key,
                    b: chartTypeKey,
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    cmbChartType.SelectedItem =
                        chartType;

                    break;
                }
            }

            txtChartYAxisTitle.Text =
                verticalAxisTitle;

            chkChartAutomaticYAxis.IsChecked =
                automaticVerticalAxis;

            txtChartYAxisMinimum.Text =
                fixedMinimum?.ToString(
                    provider: CultureInfo.InvariantCulture)
                ?? string.Empty;

            txtChartYAxisMaximum.Text =
                fixedMaximum?.ToString(
                    provider: CultureInfo.InvariantCulture)
                ?? string.Empty;

            chkChartShowLegend.IsChecked =
                showLegend;

            SelectLegendPosition(
                legendPosition: legendPosition);

            txtChartPngWidth.Text =
                pngWidth.ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartPngHeight.Text =
                pngHeight.ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartWidthMm.Text =
                widthMm.ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartHeightMm.Text =
                heightMm.ToString(
                    provider: CultureInfo.InvariantCulture);

            SelectChartResolutionDpi(
                resolutionDpi: resolutionDpi);

            SelectChartStartDateMode(
                startDateMode: startDateMode);

            UpdateChartPixelDimensions();

            dpChartStartDate.SelectedDate =
                absoluteStart.Date;

            dpChartEndDate.SelectedDate =
                absoluteEndExclusive.Date.AddDays(
                    value: -1);

            txtChartRecentDays.Text =
                Math.Max(
                    val1: 1,
                    val2:
                        (dpChartEndDate.SelectedDate.Value -
                         dpChartStartDate.SelectedDate.Value).Days)
                    .ToString(
                        provider: CultureInfo.InvariantCulture);

            txtChartTitleOverride.Text =
                chartTitle;

            #endregion


            #region Load Trigger Levels

            await LoadChartTriggerLevelsAsync(
                chartDefinitionId:
                    chartDefinitionId,
                databaseConnection:
                    databaseConnection);

            #endregion


            #region Load Series

            await LoadChartSeriesAsync(
                chartDefinitionId:
                    chartDefinitionId,
                databaseConnection:
                    databaseConnection);

            #endregion
        }


        private void SelectLegendPosition(
            string legendPosition)
        {
            for (int index = 0;
                 index < cmbChartLegendPosition.Items.Count;
                 index++)
            {
                if (cmbChartLegendPosition.Items[index]
                    is ComboBoxItem item &&
                    string.Equals(
                        a: item.Content?.ToString(),
                        b: legendPosition,
                        comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    cmbChartLegendPosition.SelectedIndex =
                        index;

                    return;
                }
            }

            cmbChartLegendPosition.SelectedIndex =
                0;
        }


        private async Task LoadChartTriggerLevelsAsync(
            int chartDefinitionId,
            SqlConnection databaseConnection)
        {
            #region Load Positive Trigger Bands

            const string sql = """
                SELECT
                    [BandName],
                    [MaximumValue]
                FROM [dbo].[ChartDefinitionTriggerBand]
                WHERE
                    [ChartDefinition_ID] = @ChartDefinition_ID
                    AND [MinimumValue] >= 0
                ORDER BY [DisplayOrder];
                """;

            decimal? green =
                null;

            decimal? amber =
                null;

            decimal? red =
                null;

            await using SqlCommand command =
                new(
                    cmdText: sql,
                    connection: databaseConnection);

            command.Parameters.Add(
                parameterName: "@ChartDefinition_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    chartDefinitionId;

            await using SqlDataReader reader =
                await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                string bandName =
                    reader.GetString(0);

                decimal maximum =
                    reader.GetDecimal(1);

                if (string.Equals(
                    a: bandName,
                    b: "Green",
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    green =
                        maximum;
                }
                else if (string.Equals(
                    a: bandName,
                    b: "Amber",
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    amber =
                        maximum;
                }
                else if (string.Equals(
                    a: bandName,
                    b: "Red",
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    red =
                        maximum;
                }
            }

            txtChartGreenTrigger.Text =
                (green ?? 2m).ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartAmberTrigger.Text =
                (amber ?? 5m).ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartRedTrigger.Text =
                (red ?? 10m).ToString(
                    provider: CultureInfo.InvariantCulture);

            #endregion
        }


        private async Task LoadChartSeriesAsync(
            int chartDefinitionId,
            SqlConnection databaseConnection)
        {
            #region Load Series Rows

            _chartSeries.Clear();

            const string sql = """
                SELECT
                    [EntityId],
                    [EntityDisplayName],
                    [DataElementKey],
                    [LegendText],
                    [ColourHex],
                    [LineWidth],
                    [MarkerSize],
                    [DisplayOrder]
                FROM [dbo].[ChartDefinitionSeries]
                WHERE [ChartDefinition_ID] = @ChartDefinition_ID
                ORDER BY [DisplayOrder];
                """;

            await using SqlCommand command =
                new(
                    cmdText: sql,
                    connection: databaseConnection);

            command.Parameters.Add(
                parameterName: "@ChartDefinition_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    chartDefinitionId;

            await using SqlDataReader reader =
                await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                _chartSeries.Add(
                    item:
                        new ChartSeriesUiItem
                        {
                            EntityId =
                                reader.GetInt32(0),

                            DisplayOrder =
                                reader.GetInt32(7),

                            EntityDisplayName =
                                reader.GetString(1),

                            DataElementDisplayName =
                                reader.GetString(2),

                            LegendText =
                                reader.GetString(3),

                            ColourHex =
                                reader.IsDBNull(4)
                                    ? string.Empty
                                    : reader.GetString(4),

                            LineWidth =
                                Convert.ToDouble(
                                    value: reader.GetDecimal(5),
                                    provider: CultureInfo.InvariantCulture),

                            MarkerSize =
                                Convert.ToDouble(
                                    value: reader.GetDecimal(6),
                                    provider: CultureInfo.InvariantCulture)
                        });
            }

            #endregion
        }


        private async Task<(int ChartDefinitionId, int ChartNumber)>
            SaveChartDefinitionAsync(
                int? chartDefinitionId,
                int? chartNumber)
        {
            #region Validate Active Project And Selected Type

            if (!_activeProjectId.HasValue)
            {
                throw new InvalidOperationException(
                    "Select an active project.");
            }

            if (cmbChartType.SelectedItem
                is not ChartTypeUiItem chartType)
            {
                throw new InvalidOperationException(
                    "Select a chart type.");
            }

            #endregion


            #region Read Editor Values

            string chartName =
                txtChartName.Text.Trim();

            DateTime absoluteStartUtc =
                DateTime.UnixEpoch;

            DateTime absoluteEndUtc =
                DateTime.UnixEpoch.AddDays(
                    value: 1);

            bool automaticVerticalAxis =
                chkChartAutomaticYAxis.IsChecked == true;

            decimal? verticalMinimum =
                automaticVerticalAxis
                    ? null
                    : decimal.Parse(
                        s: txtChartYAxisMinimum.Text,
                        style: NumberStyles.Float,
                        provider: CultureInfo.InvariantCulture);

            decimal? verticalMaximum =
                automaticVerticalAxis
                    ? null
                    : decimal.Parse(
                        s: txtChartYAxisMaximum.Text,
                        style: NumberStyles.Float,
                        provider: CultureInfo.InvariantCulture);

            decimal greenTrigger =
                decimal.Parse(
                    s: txtChartGreenTrigger.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            decimal amberTrigger =
                decimal.Parse(
                    s: txtChartAmberTrigger.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            decimal redTrigger =
                decimal.Parse(
                    s: txtChartRedTrigger.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            int pngWidth =
                int.Parse(
                    s: txtChartPngWidth.Text,
                    provider: CultureInfo.InvariantCulture);

            int pngHeight =
                int.Parse(
                    s: txtChartPngHeight.Text,
                    provider: CultureInfo.InvariantCulture);

            string legendPosition =
                GetSelectedLegendPosition();

            #endregion


            #region Open Transaction

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            using SqlTransaction transaction =
                databaseConnection.BeginTransaction(
                    iso:
                        System.Data.IsolationLevel.Serializable);

            #endregion


            try
            {
                #region Resolve Chart Type ID

                const string chartTypeSql = """
                    SELECT [ChartType_ID]
                    FROM [dbo].[ChartType] WITH (UPDLOCK, HOLDLOCK)
                    WHERE [ChartTypeKey] = @ChartTypeKey;
                    """;

                await using SqlCommand chartTypeCommand =
                    new(
                        cmdText: chartTypeSql,
                        connection: databaseConnection,
                        transaction: transaction);

                chartTypeCommand.Parameters.Add(
                    parameterName: "@ChartTypeKey",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 100)
                    .Value =
                        chartType.Key;

                object? chartTypeIdValue =
                    await chartTypeCommand.ExecuteScalarAsync();

                if (chartTypeIdValue is null ||
                    chartTypeIdValue == DBNull.Value)
                {
                    throw new InvalidOperationException(
                        "Chart type is not registered in the database.");
                }

                int chartTypeId =
                    Convert.ToInt32(
                        value: chartTypeIdValue,
                        provider: CultureInfo.InvariantCulture);

                #endregion


                #region Resolve Chart Number

                int resolvedChartNumber;

                if (chartDefinitionId.HasValue &&
                    chartNumber.HasValue)
                {
                    resolvedChartNumber =
                        chartNumber.Value;
                }
                else
                {
                    const string nextNumberSql = """
                        SELECT ISNULL(MAX([ChartNumber]), 0) + 1
                        FROM [dbo].[ChartDefinition] WITH (UPDLOCK, HOLDLOCK)
                        WHERE [Project_ID] = @Project_ID;
                        """;

                    await using SqlCommand numberCommand =
                        new(
                            cmdText: nextNumberSql,
                            connection: databaseConnection,
                            transaction: transaction);

                    numberCommand.Parameters.Add(
                        parameterName: "@Project_ID",
                        sqlDbType: System.Data.SqlDbType.Int)
                        .Value =
                            _activeProjectId.Value;

                    resolvedChartNumber =
                        Convert.ToInt32(
                            value:
                                await numberCommand.ExecuteScalarAsync(),
                            provider:
                                CultureInfo.InvariantCulture);
                }

                #endregion


                #region Insert Or Update Chart Definition

                int resolvedChartDefinitionId;

                if (chartDefinitionId.HasValue)
                {
                    const string updateSql = """
                        UPDATE [dbo].[ChartDefinition]
                        SET
                            [ChartName] = @ChartName,
                            [ChartType_ID] = @ChartType_ID,
                            [AutoTitleTemplate] = @AutoTitleTemplate,
                            [TitleOverride] = @TitleOverride,
                            [YAxisTitle] = @YAxisTitle,
                            [YAxisUnit] = @YAxisUnit,
                            [XAxisTitle] = N'Date / Time',
                            [DisplayTimeZoneId] = N'Local Time',
                            [UseAutomaticYAxis] = @UseAutomaticYAxis,
                            [FixedYAxisMinimum] = @FixedYAxisMinimum,
                            [FixedYAxisMaximum] = @FixedYAxisMaximum,
                            [ReferenceLineValue] = 0,
                            [ShowLegend] = @ShowLegend,
                            [LegendPosition] = @LegendPosition,
                            [PngWidthPixels] = @PngWidthPixels,
                            [PngHeightPixels] = @PngHeightPixels,
                            [WidthMm] = @WidthMm,
                            [HeightMm] = @HeightMm,
                            [ResolutionDpi] = @ResolutionDpi,
                            [ChartFontFamily] = N'Arial',
                            [StartDateMode] = @StartDateMode,
                            [ChartTimeWindowMode] = N'Absolute',
                            [AbsoluteStartUtc] = @AbsoluteStartUtc,
                            [AbsoluteEndUtc] = @AbsoluteEndUtc,
                            [RelativeAnchor] = NULL,
                            [RelativeStartOffsetSec] = NULL,
                            [RelativeEndOffsetSec] = NULL,
                            [ReferenceMode] = N'None',
                            [ReferenceDateTimeUtc] = NULL,
                            [ReferenceBlockSeconds] = NULL,
                            [UseMeanReference] = 1,
                            [IsEnabled] = 1,
                            [ChartOrder] = @ChartOrder,
                            [UpdatedUtc] = SYSUTCDATETIME()
                        WHERE
                            [ChartDefinition_ID] = @ChartDefinition_ID
                            AND [Project_ID] = @Project_ID
                            AND [IsDeleted] = 0;
                        """;

                    await using SqlCommand updateCommand =
                        new(
                            cmdText: updateSql,
                            connection: databaseConnection,
                            transaction: transaction);

                    AddChartDefinitionParameters(
                        command: updateCommand,
                        chartName: chartName,
                        chartTypeId: chartTypeId,
                        chartType: chartType,
                        automaticVerticalAxis: automaticVerticalAxis,
                        verticalMinimum: verticalMinimum,
                        verticalMaximum: verticalMaximum,
                        legendPosition: legendPosition,
                        absoluteStartUtc: absoluteStartUtc,
                        absoluteEndUtc: absoluteEndUtc,
                        pngWidth: pngWidth,
                        pngHeight: pngHeight,
                        chartOrder: resolvedChartNumber);

                    updateCommand.Parameters.Add(
                        parameterName: "@ChartDefinition_ID",
                        sqlDbType: System.Data.SqlDbType.Int)
                        .Value =
                            chartDefinitionId.Value;

                    int updatedRows =
                        await updateCommand.ExecuteNonQueryAsync();

                    if (updatedRows != 1)
                    {
                        throw new InvalidOperationException(
                            "Chart update failed.");
                    }

                    resolvedChartDefinitionId =
                        chartDefinitionId.Value;
                }
                else
                {
                    const string insertSql = """
                        INSERT INTO [dbo].[ChartDefinition]
                        (
                            [Project_ID],
                            [ChartNumber],
                            [ChartName],
                            [ChartType_ID],
                            [AutoTitleTemplate],
                            [TitleOverride],
                            [YAxisTitle],
                            [YAxisUnit],
                            [XAxisTitle],
                            [DisplayTimeZoneId],
                            [UseAutomaticYAxis],
                            [FixedYAxisMinimum],
                            [FixedYAxisMaximum],
                            [ReferenceLineValue],
                            [ShowLegend],
                            [LegendPosition],
                            [PngWidthPixels],
                            [PngHeightPixels],
                            [WidthMm],
                            [HeightMm],
                            [ResolutionDpi],
                            [ChartFontFamily],
                            [StartDateMode],
                            [ChartTimeWindowMode],
                            [AbsoluteStartUtc],
                            [AbsoluteEndUtc],
                            [ReferenceMode],
                            [UseMeanReference],
                            [IsEnabled],
                            [ChartOrder],
                            [IsDeleted]
                        )
                        OUTPUT INSERTED.[ChartDefinition_ID]
                        VALUES
                        (
                            @Project_ID,
                            @ChartNumber,
                            @ChartName,
                            @ChartType_ID,
                            @AutoTitleTemplate,
                            @TitleOverride,
                            @YAxisTitle,
                            @YAxisUnit,
                            N'Date / Time',
                            N'Local Time',
                            @UseAutomaticYAxis,
                            @FixedYAxisMinimum,
                            @FixedYAxisMaximum,
                            0,
                            @ShowLegend,
                            @LegendPosition,
                            @PngWidthPixels,
                            @PngHeightPixels,
                            @WidthMm,
                            @HeightMm,
                            @ResolutionDpi,
                            N'Arial',
                            @StartDateMode,
                            N'Absolute',
                            @AbsoluteStartUtc,
                            @AbsoluteEndUtc,
                            N'None',
                            1,
                            1,
                            @ChartOrder,
                            0
                        );
                        """;

                    await using SqlCommand insertCommand =
                        new(
                            cmdText: insertSql,
                            connection: databaseConnection,
                            transaction: transaction);

                    AddChartDefinitionParameters(
                        command: insertCommand,
                        chartName: chartName,
                        chartTypeId: chartTypeId,
                        chartType: chartType,
                        automaticVerticalAxis: automaticVerticalAxis,
                        verticalMinimum: verticalMinimum,
                        verticalMaximum: verticalMaximum,
                        legendPosition: legendPosition,
                        absoluteStartUtc: absoluteStartUtc,
                        absoluteEndUtc: absoluteEndUtc,
                        pngWidth: pngWidth,
                        pngHeight: pngHeight,
                        chartOrder: resolvedChartNumber);

                    insertCommand.Parameters.Add(
                        parameterName: "@ChartNumber",
                        sqlDbType: System.Data.SqlDbType.Int)
                        .Value =
                            resolvedChartNumber;

                    object? newIdValue =
                        await insertCommand.ExecuteScalarAsync();

                    if (newIdValue is null ||
                        newIdValue == DBNull.Value)
                    {
                        throw new InvalidOperationException(
                            "Chart insert failed.");
                    }

                    resolvedChartDefinitionId =
                        Convert.ToInt32(
                            value: newIdValue,
                            provider: CultureInfo.InvariantCulture);
                }

                #endregion


                #region Replace Trigger Bands

                await ReplaceChartTriggerBandsAsync(
                    chartDefinitionId:
                        resolvedChartDefinitionId,
                    greenTrigger:
                        greenTrigger,
                    amberTrigger:
                        amberTrigger,
                    redTrigger:
                        redTrigger,
                    verticalMinimum:
                        verticalMinimum,
                    verticalMaximum:
                        verticalMaximum,
                    databaseConnection:
                        databaseConnection,
                    transaction:
                        transaction);

                #endregion


                #region Replace Series

                await ReplaceChartSeriesAsync(
                    chartDefinitionId:
                        resolvedChartDefinitionId,
                    databaseConnection:
                        databaseConnection,
                    transaction:
                        transaction);

                #endregion


                #region Commit Transaction

                transaction.Commit();

                return
                (
                    ChartDefinitionId:
                        resolvedChartDefinitionId,

                    ChartNumber:
                        resolvedChartNumber
                );

                #endregion
            }
            catch
            {
                try
                {
                    transaction.Rollback();
                }
                catch
                {
                    // Preserve original exception.
                }

                throw;
            }
        }


        private void AddChartDefinitionParameters(
            SqlCommand command,
            string chartName,
            int chartTypeId,
            ChartTypeUiItem chartType,
            bool automaticVerticalAxis,
            decimal? verticalMinimum,
            decimal? verticalMaximum,
            string legendPosition,
            DateTime absoluteStartUtc,
            DateTime absoluteEndUtc,
            int pngWidth,
            int pngHeight,
            int chartOrder)
        {
            command.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    _activeProjectId
                    ?? throw new InvalidOperationException(
                        "Select an active project.");

            command.Parameters.Add(
                parameterName: "@ChartName",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 200)
                .Value =
                    chartName;

            command.Parameters.Add(
                parameterName: "@ChartType_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    chartTypeId;

            command.Parameters.Add(
                parameterName: "@AutoTitleTemplate",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 500)
                .Value =
                    "{ChartType} - {Sensors}";

            command.Parameters.Add(
                parameterName: "@TitleOverride",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 500)
                .Value =
                    string.IsNullOrWhiteSpace(
                        value: txtChartTitleOverride.Text)
                        ? DBNull.Value
                        : txtChartTitleOverride.Text.Trim();

            command.Parameters.Add(
                parameterName: "@YAxisTitle",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 200)
                .Value =
                    string.IsNullOrWhiteSpace(
                        value: txtChartYAxisTitle.Text)
                        ? chartType.VerticalAxisTitle
                        : txtChartYAxisTitle.Text.Trim();

            command.Parameters.Add(
                parameterName: "@YAxisUnit",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 50)
                .Value =
                    chartType.Unit;

            command.Parameters.Add(
                parameterName: "@UseAutomaticYAxis",
                sqlDbType: System.Data.SqlDbType.Bit)
                .Value =
                    automaticVerticalAxis;

            command.Parameters.Add(
                parameterName: "@FixedYAxisMinimum",
                sqlDbType: System.Data.SqlDbType.Decimal)
                .Value =
                    verticalMinimum.HasValue
                        ? verticalMinimum.Value
                        : DBNull.Value;

            command.Parameters["@FixedYAxisMinimum"].Precision =
                18;

            command.Parameters["@FixedYAxisMinimum"].Scale =
                6;

            command.Parameters.Add(
                parameterName: "@FixedYAxisMaximum",
                sqlDbType: System.Data.SqlDbType.Decimal)
                .Value =
                    verticalMaximum.HasValue
                        ? verticalMaximum.Value
                        : DBNull.Value;

            command.Parameters["@FixedYAxisMaximum"].Precision =
                18;

            command.Parameters["@FixedYAxisMaximum"].Scale =
                6;

            command.Parameters.Add(
                parameterName: "@ShowLegend",
                sqlDbType: System.Data.SqlDbType.Bit)
                .Value =
                    chkChartShowLegend.IsChecked == true;

            command.Parameters.Add(
                parameterName: "@LegendPosition",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 50)
                .Value =
                    legendPosition;

            command.Parameters.Add(
                parameterName: "@PngWidthPixels",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    pngWidth;

            command.Parameters.Add(
                parameterName: "@PngHeightPixels",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    pngHeight;

            command.Parameters.Add(
                parameterName: "@WidthMm",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    int.Parse(
                        s: txtChartWidthMm.Text,
                        provider: CultureInfo.InvariantCulture);

            command.Parameters.Add(
                parameterName: "@HeightMm",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    int.Parse(
                        s: txtChartHeightMm.Text,
                        provider: CultureInfo.InvariantCulture);

            command.Parameters.Add(
                parameterName: "@ResolutionDpi",
                sqlDbType: System.Data.SqlDbType.SmallInt)
                .Value =
                    GetSelectedChartResolutionDpi();

            command.Parameters.Add(
                parameterName: "@StartDateMode",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 20)
                .Value =
                    GetSelectedChartStartDateMode();

            command.Parameters.Add(
                parameterName: "@AbsoluteStartUtc",
                sqlDbType: System.Data.SqlDbType.DateTime2)
                .Value =
                    absoluteStartUtc;

            command.Parameters.Add(
                parameterName: "@AbsoluteEndUtc",
                sqlDbType: System.Data.SqlDbType.DateTime2)
                .Value =
                    absoluteEndUtc;

            command.Parameters.Add(
                parameterName: "@ChartOrder",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    chartOrder;
        }


        private string GetSelectedLegendPosition()
        {
            if (cmbChartLegendPosition.SelectedItem
                is ComboBoxItem item)
            {
                return item.Content?.ToString()?.Trim()
                    ?? "Below Title";
            }

            return "Below Title";
        }


        private async Task ReplaceChartTriggerBandsAsync(
            int chartDefinitionId,
            decimal greenTrigger,
            decimal amberTrigger,
            decimal redTrigger,
            decimal? verticalMinimum,
            decimal? verticalMaximum,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Delete Existing Bands

            const string deleteSql = """
                DELETE FROM [dbo].[ChartDefinitionTriggerBand]
                WHERE [ChartDefinition_ID] = @ChartDefinition_ID;
                """;

            await using (SqlCommand deleteCommand =
                new(
                    cmdText: deleteSql,
                    connection: databaseConnection,
                    transaction: transaction))
            {
                deleteCommand.Parameters.Add(
                    parameterName: "@ChartDefinition_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value =
                        chartDefinitionId;

                await deleteCommand.ExecuteNonQueryAsync();
            }

            #endregion


            #region Determine Band Extent

            decimal positiveExtent =
                verticalMaximum.HasValue
                    ? Math.Max(
                        redTrigger,
                        verticalMaximum.Value)
                    : redTrigger * 2m;

            decimal negativeExtent =
                verticalMinimum.HasValue
                    ? Math.Min(
                        -redTrigger,
                        verticalMinimum.Value)
                    : -positiveExtent;

            #endregion


            #region Insert Mirrored Bands

            List<(string Name, decimal Minimum, decimal Maximum, string Colour, int Order)>
                bands =
                    new()
                    {
                        ("Grey", negativeExtent, -redTrigger, ChartAboveRedBandColourHex, 1),
                        ("Red", -redTrigger, -amberTrigger, ChartRedBandColourHex, 2),
                        ("Amber", -amberTrigger, -greenTrigger, ChartAmberBandColourHex, 3),
                        ("Green", -greenTrigger, 0m, ChartGreenBandColourHex, 4),
                        ("Green", 0m, greenTrigger, ChartGreenBandColourHex, 5),
                        ("Amber", greenTrigger, amberTrigger, ChartAmberBandColourHex, 6),
                        ("Red", amberTrigger, redTrigger, ChartRedBandColourHex, 7),
                        ("Grey", redTrigger, positiveExtent, ChartAboveRedBandColourHex, 8)
                    };

            const string insertSql = """
                INSERT INTO [dbo].[ChartDefinitionTriggerBand]
                (
                    [ChartDefinition_ID],
                    [BandName],
                    [MinimumValue],
                    [MaximumValue],
                    [IsSymmetric],
                    [ColourHex],
                    [Opacity],
                    [DrawBoundaryLines],
                    [DisplayOrder]
                )
                VALUES
                (
                    @ChartDefinition_ID,
                    @BandName,
                    @MinimumValue,
                    @MaximumValue,
                    0,
                    @ColourHex,
                    1.0,
                    0,
                    @DisplayOrder
                );
                """;

            foreach ((string Name, decimal Minimum, decimal Maximum, string Colour, int Order)
                band in bands)
            {
                if (band.Minimum >= band.Maximum)
                {
                    continue;
                }

                await using SqlCommand insertCommand =
                    new(
                        cmdText: insertSql,
                        connection: databaseConnection,
                        transaction: transaction);

                insertCommand.Parameters.Add(
                    parameterName: "@ChartDefinition_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value =
                        chartDefinitionId;

                insertCommand.Parameters.Add(
                    parameterName: "@BandName",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 100)
                    .Value =
                        band.Name;

                SqlParameter minimumParameter =
                    insertCommand.Parameters.Add(
                        parameterName: "@MinimumValue",
                        sqlDbType: System.Data.SqlDbType.Decimal);

                minimumParameter.Precision =
                    18;

                minimumParameter.Scale =
                    6;

                minimumParameter.Value =
                    band.Minimum;

                SqlParameter maximumParameter =
                    insertCommand.Parameters.Add(
                        parameterName: "@MaximumValue",
                        sqlDbType: System.Data.SqlDbType.Decimal);

                maximumParameter.Precision =
                    18;

                maximumParameter.Scale =
                    6;

                maximumParameter.Value =
                    band.Maximum;

                insertCommand.Parameters.Add(
                    parameterName: "@ColourHex",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 20)
                    .Value =
                        band.Colour;

                insertCommand.Parameters.Add(
                    parameterName: "@DisplayOrder",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value =
                        band.Order;

                await insertCommand.ExecuteNonQueryAsync();
            }

            #endregion
        }


        private async Task ReplaceChartSeriesAsync(
            int chartDefinitionId,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Delete Existing Series

            const string deleteSql = """
                DELETE FROM [dbo].[ChartDefinitionSeries]
                WHERE [ChartDefinition_ID] = @ChartDefinition_ID;
                """;

            await using (SqlCommand deleteCommand =
                new(
                    cmdText: deleteSql,
                    connection: databaseConnection,
                    transaction: transaction))
            {
                deleteCommand.Parameters.Add(
                    parameterName: "@ChartDefinition_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value =
                        chartDefinitionId;

                await deleteCommand.ExecuteNonQueryAsync();
            }

            #endregion


            #region Insert Configured Series

            const string insertSql = """
                INSERT INTO [dbo].[ChartDefinitionSeries]
                (
                    [ChartDefinition_ID],
                    [EntityId],
                    [EntityDisplayName],
                    [DataElementKey],
                    [LegendText],
                    [ColourHex],
                    [LineWidth],
                    [MarkerSize],
                    [DisplayOrder]
                )
                VALUES
                (
                    @ChartDefinition_ID,
                    @EntityId,
                    @EntityDisplayName,
                    @DataElementKey,
                    @LegendText,
                    NULLIF(@ColourHex, N''),
                    @LineWidth,
                    @MarkerSize,
                    @DisplayOrder
                );
                """;

            foreach (ChartSeriesUiItem series
                in _chartSeries.OrderBy(
                    keySelector: item => item.DisplayOrder))
            {
                if (series.EntityId <= 0)
                {
                    throw new InvalidOperationException(
                        $"Series {series.DisplayOrder} does not reference a valid database entity.");
                }

                await using SqlCommand insertCommand =
                    new(
                        cmdText: insertSql,
                        connection: databaseConnection,
                        transaction: transaction);

                insertCommand.Parameters.Add(
                    parameterName: "@ChartDefinition_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value = chartDefinitionId;

                insertCommand.Parameters.Add(
                    parameterName: "@EntityId",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value = series.EntityId;

                insertCommand.Parameters.Add(
                    parameterName: "@EntityDisplayName",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 200)
                    .Value = series.EntityDisplayName;

                insertCommand.Parameters.Add(
                    parameterName: "@DataElementKey",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 100)
                    .Value = series.DataElementDisplayName;

                insertCommand.Parameters.Add(
                    parameterName: "@LegendText",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 300)
                    .Value = series.LegendText;

                insertCommand.Parameters.Add(
                    parameterName: "@ColourHex",
                    sqlDbType: System.Data.SqlDbType.NVarChar,
                    size: 20)
                    .Value = series.ColourHex;

                AddSeriesDecimalParameter(
                    command: insertCommand,
                    parameterName: "@LineWidth",
                    value: series.LineWidth);

                AddSeriesDecimalParameter(
                    command: insertCommand,
                    parameterName: "@MarkerSize",
                    value: series.MarkerSize);

                insertCommand.Parameters.Add(
                    parameterName: "@DisplayOrder",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value = series.DisplayOrder;

                await insertCommand.ExecuteNonQueryAsync();
            }

            #endregion
        }


        private static void AddSeriesDecimalParameter(
            SqlCommand command,
            string parameterName,
            double value)
        {
            SqlParameter parameter =
                command.Parameters.Add(
                    parameterName: parameterName,
                    sqlDbType: System.Data.SqlDbType.Decimal);

            parameter.Precision =
                10;

            parameter.Scale =
                3;

            parameter.Value =
                Convert.ToDecimal(
                    value: value,
                    provider: CultureInfo.InvariantCulture);
        }


        private async Task SoftDeleteChartDefinitionAsync(
            int chartDefinitionId)
        {
            #region Soft Delete Chart

            const string sql = """
                UPDATE [dbo].[ChartDefinition]
                SET
                    [IsDeleted] = 1,
                    [UpdatedUtc] = SYSUTCDATETIME()
                WHERE
                    [ChartDefinition_ID] = @ChartDefinition_ID
                    AND [Project_ID] = @Project_ID
                    AND [IsDeleted] = 0;
                """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand command =
                new(
                    cmdText: sql,
                    connection: databaseConnection);

            command.Parameters.Add(
                parameterName: "@ChartDefinition_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    chartDefinitionId;

            command.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    _activeProjectId
                    ?? throw new InvalidOperationException(
                        "Select an active project.");

            int affectedRows =
                await command.ExecuteNonQueryAsync();

            if (affectedRows != 1)
            {
                throw new InvalidOperationException(
                    "Chart delete failed.");
            }

            #endregion
        }

        #endregion


        #region Chart Preview Data

        private async Task LoadChartPreviewDataAsync()
        {
            if (!_activeProjectId.HasValue ||
                cmbChartType.SelectedItem is not ChartTypeUiItem chartType)
            {
                throw new InvalidOperationException(
                    "Select an active project and chart type.");
            }

            const string projectSql = """
                SELECT [ProjectStartDate], [TimeZoneId]
                FROM [dbo].[Project]
                WHERE [Project_ID] = @Project_ID AND [IsDeleted] = 0;
                """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString: GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            DateTime projectStartDate;
            string timeZoneId;

            await using (SqlCommand projectCommand =
                new(
                    cmdText: projectSql,
                    connection: databaseConnection))
            {
                projectCommand.Parameters.Add(
                    parameterName: "@Project_ID",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value = _activeProjectId.Value;

                await using SqlDataReader projectReader =
                    await projectCommand.ExecuteReaderAsync();

                if (!await projectReader.ReadAsync())
                {
                    throw new InvalidOperationException(
                        "The active project no longer exists.");
                }

                projectStartDate = projectReader.GetDateTime(0).Date;
                timeZoneId = projectReader.IsDBNull(1)
                    ? TimeZoneInfo.Utc.Id
                    : projectReader.GetString(1);
            }

            TimeZoneInfo projectTimeZone =
                TimeZoneInfo.FindSystemTimeZoneById(
                    id: timeZoneId);

            DateTime projectToday =
                TimeZoneInfo.ConvertTimeFromUtc(
                    dateTime: DateTime.UtcNow,
                    destinationTimeZone: projectTimeZone)
                .Date;

            DateTime previewStartDate =
                string.Equals(
                    a: GetSelectedChartStartDateMode(),
                    b: ChartStartDateModeProjectStart,
                    comparisonType: StringComparison.Ordinal)
                    ? projectStartDate
                    : projectToday.AddDays(
                        value: -14);

            DateTime previewEndExclusiveDate =
                projectToday.AddDays(
                    value: 1);

            _chartPreviewStartUtc = TimeZoneInfo.ConvertTimeToUtc(
                dateTime: DateTime.SpecifyKind(
                    value: previewStartDate,
                    kind: DateTimeKind.Unspecified),
                sourceTimeZone: projectTimeZone);

            _chartPreviewEndUtcExclusive = TimeZoneInfo.ConvertTimeToUtc(
                dateTime: DateTime.SpecifyKind(
                    value: previewEndExclusiveDate,
                    kind: DateTimeKind.Unspecified),
                sourceTimeZone: projectTimeZone);

            dpChartStartDate.SelectedDate = previewStartDate;
            dpChartEndDate.SelectedDate = projectToday;

            _chartPreviewSeries.Clear();

            Brush[] palette =
            [
                Brushes.Blue,
                Brushes.DarkOrange,
                Brushes.ForestGreen,
                Brushes.Purple,
                Brushes.Crimson,
                Brushes.Teal,
                Brushes.SaddleBrown,
                Brushes.DeepPink
            ];

            int seriesIndex = 0;

            foreach (ChartSeriesUiItem configuredSeries
                in _chartSeries.OrderBy(
                    keySelector: series => series.DisplayOrder))
            {
                string valueExpression =
                    GetChartPreviewValueExpression(
                        chartType: chartType,
                        dataElement: configuredSeries.DataElementDisplayName);

                string entityIdColumn =
                    GetChartEntityIdColumnName(
                        entityKind: chartType.EntityKind);

                string joins =
                    chartType.Key switch
                    {
                        "Displacement" =>
                            " INNER JOIN [dbo].[CoordinatesReference] AS CR ON CR.[PointName_ID] = E.[PointName_ID] AND CR.[IsDeleted] = 0 ",
                        "ToRLevel" =>
                            " INNER JOIN [dbo].[ToRReference] AS TR ON TR.[PointName_ID] = E.[PointName_ID] AND TR.[IsDeleted] = 0 ",
                        _ => string.Empty
                    };

                string seriesSql =
                    $"SELECT E.[UTCtime], {valueExpression} AS [PreviewValue] " +
                    $"FROM [dbo].[{chartType.SourceTable}] AS E {joins}" +
                    $"WHERE E.[{entityIdColumn}] = @EntityId " +
                    "AND E.[UTCtime] >= @StartUtc AND E.[UTCtime] < @EndUtcExclusive " +
                    "AND E.[IsDeleted] = 0 ORDER BY E.[UTCtime];";

                ChartPreviewSeries previewSeries =
                    new()
                    {
                        LegendText = configuredSeries.LegendText,
                        Stroke = palette[seriesIndex % palette.Length],
                        LineWidth = Math.Max(
                            val1: 0.5,
                            val2: configuredSeries.LineWidth)
                    };

                await using SqlCommand seriesCommand =
                    new(
                        cmdText: seriesSql,
                        connection: databaseConnection);

                seriesCommand.Parameters.Add(
                    parameterName: "@EntityId",
                    sqlDbType: System.Data.SqlDbType.Int)
                    .Value = configuredSeries.EntityId;

                seriesCommand.Parameters.Add(
                    parameterName: "@StartUtc",
                    sqlDbType: System.Data.SqlDbType.DateTime2)
                    .Value = _chartPreviewStartUtc;

                seriesCommand.Parameters.Add(
                    parameterName: "@EndUtcExclusive",
                    sqlDbType: System.Data.SqlDbType.DateTime2)
                    .Value = _chartPreviewEndUtcExclusive;

                await using SqlDataReader seriesReader =
                    await seriesCommand.ExecuteReaderAsync();

                while (await seriesReader.ReadAsync())
                {
                    if (!seriesReader.IsDBNull(1))
                    {
                        previewSeries.Points.Add(
                            item:
                                new ChartPreviewPoint(
                                    UtcTime: seriesReader.GetDateTime(0),
                                    Value: Convert.ToDouble(
                                        value: seriesReader.GetValue(1),
                                        provider: CultureInfo.InvariantCulture)));
                    }
                }

                _chartPreviewSeries.Add(
                    item: previewSeries);

                seriesIndex++;
            }
        }


        private static string GetChartPreviewValueExpression(
            ChartTypeUiItem chartType,
            string dataElement)
        {
            return (chartType.Key, dataElement) switch
            {
                ("Displacement", "dE") => "(E.[E] - CR.[Eref]) * 1000.0",
                ("Displacement", "dN") => "(E.[N] - CR.[Nref]) * 1000.0",
                ("Displacement", "dH") => "(E.[H] - CR.[Href]) * 1000.0",
                ("Displacement", "d2D") => "SQRT(POWER(E.[E] - CR.[Eref], 2) + POWER(E.[N] - CR.[Nref], 2)) * 1000.0",
                ("Displacement", "d3D") => "SQRT(POWER(E.[E] - CR.[Eref], 2) + POWER(E.[N] - CR.[Nref], 2) + POWER(E.[H] - CR.[Href], 2)) * 1000.0",
                ("ToRLevel", "ToR - Reference ToR") => "(E.[ToR] - TR.[ToR]) * 1000.0",
                ("StructuralArray", "d2D") or ("CrownData", "d2D") => "SQRT(POWER(E.[dE], 2) + POWER(E.[dN], 2)) * 1000.0",
                ("StructuralArray", "d3D") or ("CrownData", "d3D") => "SQRT(POWER(E.[dE], 2) + POWER(E.[dN], 2) + POWER(E.[dH], 2)) * 1000.0",
                ("StructuralArray", "dE") or ("StructuralArray", "dN") or ("StructuralArray", "dH") or
                ("CrownData", "dE") or ("CrownData", "dN") or ("CrownData", "dH") => $"E.[{dataElement}] * 1000.0",
                ("PrismTiltMmPerM", "TiltX") => "E.[TiltX_MperM] * 1000.0",
                ("PrismTiltMmPerM", "TiltY") => "E.[TiltY_MperM] * 1000.0",
                ("CrackMeter", "d2D") or ("CrackMeter", "d3D") or ("CrackMeter", "dH") => $"E.[{dataElement}] * 1000.0",
                _ when chartType.DataElements.Contains(
                    value: dataElement,
                    comparer: StringComparer.Ordinal) => $"E.[{dataElement}]",
                _ => throw new InvalidOperationException(
                    $"Unsupported preview data element '{dataElement}' for '{chartType.DisplayName}'.")
            };
        }

        #endregion


        private async void btnChartPreview_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Canvas Configuration

            if (!TryValidateChartCanvasConfiguration(
                validationMessage: out string validationMessage))
            {
                txtChartStatus.Text =
                    validationMessage;

                return;
            }

            #endregion


            #region Load Actual Preview Measurements

            try
            {
                await LoadChartPreviewDataAsync();
            }
            catch (Exception ex)
            {
                txtChartStatus.Text =
                    $"Unable to load chart preview data: {ex.Message}";

                return;
            }

            #endregion


            #region Display Final-Size Preview Flyout

            int widthMm =
                int.Parse(
                    s: txtChartWidthMm.Text,
                    provider: CultureInfo.InvariantCulture);

            int heightMm =
                int.Parse(
                    s: txtChartHeightMm.Text,
                    provider: CultureInfo.InvariantCulture);

            cnvChartPreview.Width =
                widthMm * 96.0 / 25.4;

            cnvChartPreview.Height =
                heightMm * 96.0 / 25.4;

            popChartPreview.IsOpen =
                true;

            cnvChartPreview.UpdateLayout();

            popChartPreview.Child.UpdateLayout();

            FrameworkElement previewContent =
                popChartPreview.Child as FrameworkElement
                ?? throw new InvalidOperationException(
                    "The chart preview content must be a FrameworkElement.");

            Rect workArea =
                SystemParameters.WorkArea;

            popChartPreview.HorizontalOffset =
                workArea.Left +
                Math.Max(
                    val1: 0,
                    val2: (workArea.Width - previewContent.ActualWidth) / 2.0);

            popChartPreview.VerticalOffset =
                workArea.Top +
                Math.Max(
                    val1: 0,
                    val2: (workArea.Height - previewContent.ActualHeight) / 2.0);

            await Dispatcher.InvokeAsync(
                callback:
                    new Action(
                        _chartPreviewRenderer.Render));

            txtChartPreviewStatus.Text =
                $"Final physical size: {widthMm} x {heightMm} mm; " +
                $"production bitmap: {txtChartCalculatedPixels.Text}; {DefaultChartFontFamily}.";

            txtChartStatus.Text =
                $"Preview generated at {txtChartCalculatedPixels.Text} using " +
                $"{DefaultChartFontFamily}. The renderer boundary is ready for ScottPlot 5.";

            #endregion
        }


        private void btnCloseChartPreview_Click(
            object sender,
            RoutedEventArgs e)
        {
            popChartPreview.IsOpen =
                false;
        }


        private bool TryValidateChartCanvasConfiguration(
            out string validationMessage)
        {
            #region Validate Chart Type

            if (cmbChartType.SelectedItem
                is not ChartTypeUiItem)
            {
                validationMessage =
                    "Definition: Select a chart type.";

                return false;
            }

            #endregion


            #region Validate Report-Resolved Time Basis

            string startDateMode =
                GetSelectedChartStartDateMode();

            if (!string.Equals(
                    a: startDateMode,
                    b: ChartStartDateModeReportStart,
                    comparisonType: StringComparison.Ordinal) &&
                !string.Equals(
                    a: startDateMode,
                    b: ChartStartDateModeProjectStart,
                    comparisonType: StringComparison.Ordinal))
            {
                validationMessage =
                    "Time Basis: Select Report Start or Project Start.";

                return false;
            }

            #endregion


            #region Validate Triggers

            if (!TryReadPositiveDecimal(
                text: txtChartGreenTrigger.Text,
                valueName: "Trigger Bands: Green trigger",
                value: out decimal greenTrigger,
                validationMessage: out validationMessage))
            {
                return false;
            }

            if (!TryReadPositiveDecimal(
                text: txtChartAmberTrigger.Text,
                valueName: "Trigger Bands: Amber trigger",
                value: out decimal amberTrigger,
                validationMessage: out validationMessage))
            {
                return false;
            }

            if (!TryReadPositiveDecimal(
                text: txtChartRedTrigger.Text,
                valueName: "Trigger Bands: Red trigger",
                value: out decimal redTrigger,
                validationMessage: out validationMessage))
            {
                return false;
            }

            if (!(greenTrigger < amberTrigger &&
                  amberTrigger < redTrigger))
            {
                validationMessage =
                    "Trigger Bands: Values must satisfy Green < Amber < Red.";

                return false;
            }

            #endregion


            #region Validate Vertical Axis

            decimal verticalMinimum;
            decimal verticalMaximum;

            if (chkChartAutomaticYAxis.IsChecked == true)
            {
                ApplyDefaultVerticalAxisFromRedTrigger();
            }

            if (!decimal.TryParse(
                s: txtChartYAxisMinimum.Text,
                style: NumberStyles.Float,
                provider: CultureInfo.InvariantCulture,
                result: out verticalMinimum) ||
                !decimal.TryParse(
                    s: txtChartYAxisMaximum.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture,
                    result: out verticalMaximum))
            {
                validationMessage =
                    "Axes & Titles: Enter valid numeric Vertical Axis limits.";

                return false;
            }

            if (verticalMinimum >= verticalMaximum)
            {
                validationMessage =
                    "Axes & Titles: Minimum must be less than Maximum.";

                return false;
            }

            if (verticalMinimum > -redTrigger ||
                verticalMaximum < redTrigger)
            {
                validationMessage =
                    "Axes & Titles: Vertical Axis limits must extend beyond both Red trigger levels.";

                return false;
            }

            #endregion


            #region Validate Appearance

            if (!TryValidateChartAppearanceConfiguration(
                validationMessage: out validationMessage))
            {
                return false;
            }

            #endregion

            validationMessage =
                string.Empty;

            return true;
        }


        #region Chart Physical Size And Renderer Boundary

        private void NumericTextBox_PreviewTextInput(
            object sender,
            TextCompositionEventArgs e)
        {
            #region Validate Typed Numeric Input

            if (sender is not TextBox textBox)
            {
                e.Handled =
                    true;

                return;
            }

            string proposedText =
                BuildProspectiveText(
                    textBox: textBox,
                    insertedText: e.Text);

            e.Handled =
                !IsValidNumericEdit(
                    proposedText: proposedText,
                    validationMode: textBox.Tag as string);

            #endregion
        }


        private void NumericTextBox_Pasting(
            object sender,
            DataObjectPastingEventArgs e)
        {
            #region Validate Pasted Numeric Input

            if (sender is not TextBox textBox ||
                !e.DataObject.GetDataPresent(
                    format: DataFormats.UnicodeText))
            {
                e.CancelCommand();

                return;
            }

            string pastedText =
                e.DataObject.GetData(
                    format: DataFormats.UnicodeText) as string
                ?? string.Empty;

            string proposedText =
                BuildProspectiveText(
                    textBox: textBox,
                    insertedText: pastedText);

            if (!IsValidNumericEdit(
                proposedText: proposedText,
                validationMode: textBox.Tag as string))
            {
                e.CancelCommand();
            }

            #endregion
        }


        private static string BuildProspectiveText(
            TextBox textBox,
            string insertedText)
        {
            #region Construct Prospective Text

            string existingText =
                textBox.Text
                ?? string.Empty;

            return existingText.Remove(
                    startIndex: textBox.SelectionStart,
                    count: textBox.SelectionLength)
                .Insert(
                    startIndex: textBox.SelectionStart,
                    value: insertedText);

            #endregion
        }


        private static bool IsValidNumericEdit(
            string proposedText,
            string? validationMode)
        {
            #region Validate Prospective Numeric Text

            if (string.IsNullOrEmpty(
                value: proposedText))
            {
                return true;
            }

            if (string.Equals(
                a: validationMode,
                b: "PositiveInteger",
                comparisonType: StringComparison.Ordinal))
            {
                return proposedText.All(
                    predicate: char.IsDigit);
            }

            if (proposedText is "-" or "." or "-.")
            {
                return true;
            }

            return decimal.TryParse(
                s: proposedText,
                style:
                    NumberStyles.AllowLeadingSign |
                    NumberStyles.AllowDecimalPoint,
                provider: CultureInfo.InvariantCulture,
                result: out _);

            #endregion
        }


        private bool TryValidateChartAppearanceConfiguration(
            out string validationMessage)
        {
            #region Validate Physical Dimensions

            if (!int.TryParse(
                s: txtChartWidthMm.Text,
                style: NumberStyles.Integer,
                provider: CultureInfo.InvariantCulture,
                result: out int widthMm) ||
                widthMm <= 10)
            {
                validationMessage =
                    "Appearance: Chart width must be a whole number greater than 10 millimetres.";

                return false;
            }

            if (!int.TryParse(
                s: txtChartHeightMm.Text,
                style: NumberStyles.Integer,
                provider: CultureInfo.InvariantCulture,
                result: out int heightMm) ||
                heightMm <= 10)
            {
                validationMessage =
                    "Appearance: Chart height must be a whole number greater than 10 millimetres.";

                return false;
            }

            int resolutionDpi =
                GetSelectedChartResolutionDpi();

            if (resolutionDpi != 300 &&
                resolutionDpi != 600)
            {
                validationMessage =
                    "Appearance: Chart resolution must be 300 or 600 DPI.";

                return false;
            }

            #endregion


            #region Validate Line And Marker Sizes

            if (!double.TryParse(
                s: txtChartDefaultLineWidth.Text,
                style: NumberStyles.Float,
                provider: CultureInfo.InvariantCulture,
                result: out double lineWidth) ||
                lineWidth <= 0)
            {
                validationMessage =
                    "Appearance: Line width must be greater than zero.";

                return false;
            }

            if (!double.TryParse(
                s: txtChartDefaultMarkerSize.Text,
                style: NumberStyles.Float,
                provider: CultureInfo.InvariantCulture,
                result: out double markerSize) ||
                markerSize <= 0)
            {
                validationMessage =
                    "Appearance: Marker size must be greater than zero.";

                return false;
            }

            #endregion

            validationMessage =
                string.Empty;

            return true;
        }

        private void ChartPhysicalSize_Changed(
            object sender,
            RoutedEventArgs e)
        {
            if (!IsInitialized)
            {
                return;
            }

            UpdateChartPixelDimensions();
        }


        private void cmbChartStartDateMode_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            if (!IsInitialized)
            {
                return;
            }

            txtChartStartDateBasis.Text =
                string.Equals(
                    a: GetSelectedChartStartDateMode(),
                    b: ChartStartDateModeProjectStart,
                    comparisonType: StringComparison.Ordinal)
                    ? "Project Start"
                    : "Report Start";
        }


        private void ChartStartDateBasis_Checked(
            object sender,
            RoutedEventArgs e)
        {
            if (!IsInitialized)
            {
                return;
            }

            SelectChartStartDateMode(
                startDateMode:
                    rbChartProjectStart.IsChecked == true
                        ? ChartStartDateModeProjectStart
                        : ChartStartDateModeReportStart);
        }


        private string GetSelectedChartStartDateMode()
        {
            return (cmbChartStartDateMode.SelectedItem as ComboBoxItem)?.Tag?.ToString()
                   ?? ChartStartDateModeReportStart;
        }


        private void SelectChartStartDateMode(
            string startDateMode)
        {
            foreach (object item in cmbChartStartDateMode.Items)
            {
                if (item is ComboBoxItem comboBoxItem &&
                    string.Equals(
                        a: comboBoxItem.Tag?.ToString(),
                        b: startDateMode,
                        comparisonType: StringComparison.Ordinal))
                {
                    cmbChartStartDateMode.SelectedItem =
                        comboBoxItem;

                    rbChartProjectStart.IsChecked =
                        string.Equals(
                            a: startDateMode,
                            b: ChartStartDateModeProjectStart,
                            comparisonType: StringComparison.Ordinal);

                    rbChartReportStart.IsChecked =
                        rbChartProjectStart.IsChecked != true;

                    return;
                }
            }

            cmbChartStartDateMode.SelectedIndex =
                0;
        }


        private int GetSelectedChartResolutionDpi()
        {
            string? value =
                (cmbChartResolutionDpi.SelectedItem as ComboBoxItem)?.Tag?.ToString();

            return int.TryParse(
                s: value,
                style: NumberStyles.Integer,
                provider: CultureInfo.InvariantCulture,
                result: out int resolutionDpi)
                ? resolutionDpi
                : DefaultChartResolutionDpi;
        }


        private void SelectChartResolutionDpi(
            int resolutionDpi)
        {
            foreach (object item in cmbChartResolutionDpi.Items)
            {
                if (item is ComboBoxItem comboBoxItem &&
                    int.TryParse(
                        s: comboBoxItem.Tag?.ToString(),
                        style: NumberStyles.Integer,
                        provider: CultureInfo.InvariantCulture,
                        result: out int itemDpi) &&
                    itemDpi == resolutionDpi)
                {
                    cmbChartResolutionDpi.SelectedItem =
                        comboBoxItem;

                    return;
                }
            }

            cmbChartResolutionDpi.SelectedIndex =
                0;
        }


        private void UpdateChartPixelDimensions()
        {
            if (!int.TryParse(
                    s: txtChartWidthMm.Text,
                    style: NumberStyles.Integer,
                    provider: CultureInfo.InvariantCulture,
                    result: out int widthMm) ||
                !int.TryParse(
                    s: txtChartHeightMm.Text,
                    style: NumberStyles.Integer,
                    provider: CultureInfo.InvariantCulture,
                    result: out int heightMm) ||
                widthMm <= 10 ||
                heightMm <= 10)
            {
                txtChartCalculatedPixels.Text =
                    "Invalid physical size";

                return;
            }

            int resolutionDpi =
                GetSelectedChartResolutionDpi();

            int pngWidth =
                (int)Math.Round(
                    d: widthMm * resolutionDpi / 25.4m,
                    mode: MidpointRounding.AwayFromZero);

            int pngHeight =
                (int)Math.Round(
                    d: pngWidth * (decimal)heightMm / widthMm,
                    mode: MidpointRounding.AwayFromZero);

            txtChartPngWidth.Text =
                pngWidth.ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartPngHeight.Text =
                pngHeight.ToString(
                    provider: CultureInfo.InvariantCulture);

            txtChartCalculatedPixels.Text =
                $"{pngWidth} x {pngHeight} px";
        }

        #endregion


        private void RenderBlankChartCanvas()
        {
            #region Resolve Canvas Size

            cnvChartPreview.Children.Clear();

            double canvasWidth =
                cnvChartPreview.ActualWidth > 100
                    ? cnvChartPreview.ActualWidth
                    : 900;

            double canvasHeight =
                cnvChartPreview.ActualHeight > 100
                    ? cnvChartPreview.ActualHeight
                    : 300;

            #endregion


            #region Resolve Engineering Values

            ChartTypeUiItem chartType =
                (ChartTypeUiItem)cmbChartType.SelectedItem!;

            decimal greenTrigger =
                decimal.Parse(
                    s: txtChartGreenTrigger.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            decimal amberTrigger =
                decimal.Parse(
                    s: txtChartAmberTrigger.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            decimal redTrigger =
                decimal.Parse(
                    s: txtChartRedTrigger.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            decimal verticalMinimum =
                decimal.Parse(
                    s: txtChartYAxisMinimum.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            decimal verticalMaximum =
                decimal.Parse(
                    s: txtChartYAxisMaximum.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture);

            #endregion


            #region Define Plot Rectangle

            const double leftMargin =
                72;

            const double rightMargin =
                24;

            const double topMargin =
                58;

            const double bottomMargin =
                92;

            double plotLeft =
                leftMargin;

            double plotTop =
                topMargin;

            double plotWidth =
                Math.Max(
                    100,
                    canvasWidth - leftMargin - rightMargin);

            double plotHeight =
                Math.Max(
                    80,
                    canvasHeight - topMargin - bottomMargin);

            double plotBottom =
                plotTop + plotHeight;

            #endregion


            #region Local Conversion Helpers

            double ToY(
                decimal engineeringValue)
            {
                decimal range =
                    verticalMaximum - verticalMinimum;

                decimal fraction =
                    (verticalMaximum - engineeringValue) /
                    range;

                return plotTop +
                       ((double)fraction * plotHeight);
            }

            Brush BrushFromHex(
                string colourHex)
            {
                return
                    (Brush)new BrushConverter().ConvertFromString(
                        colourHex)!;
            }

            void AddBand(
                decimal minimum,
                decimal maximum,
                string colourHex)
            {
                decimal clippedMinimum =
                    Math.Max(
                        minimum,
                        verticalMinimum);

                decimal clippedMaximum =
                    Math.Min(
                        maximum,
                        verticalMaximum);

                if (clippedMinimum >= clippedMaximum)
                {
                    return;
                }

                double yTop =
                    ToY(
                        engineeringValue:
                            clippedMaximum);

                double yBottom =
                    ToY(
                        engineeringValue:
                            clippedMinimum);

                Rectangle rectangle =
                    new()
                    {
                        Width =
                            plotWidth,

                        Height =
                            yBottom - yTop,

                        Fill =
                            BrushFromHex(
                                colourHex:
                                    colourHex),

                        StrokeThickness =
                            0
                    };

                Canvas.SetLeft(
                    element: rectangle,
                    length: plotLeft);

                Canvas.SetTop(
                    element: rectangle,
                    length: yTop);

                cnvChartPreview.Children.Add(
                    element: rectangle);
            }

            #endregion


            #region Draw Trigger Backdrop

            AddBand(
                minimum: verticalMinimum,
                maximum: -redTrigger,
                colourHex: ChartAboveRedBandColourHex);

            AddBand(
                minimum: -redTrigger,
                maximum: -amberTrigger,
                colourHex: ChartRedBandColourHex);

            AddBand(
                minimum: -amberTrigger,
                maximum: -greenTrigger,
                colourHex: ChartAmberBandColourHex);

            AddBand(
                minimum: -greenTrigger,
                maximum: greenTrigger,
                colourHex: ChartGreenBandColourHex);

            AddBand(
                minimum: greenTrigger,
                maximum: amberTrigger,
                colourHex: ChartAmberBandColourHex);

            AddBand(
                minimum: amberTrigger,
                maximum: redTrigger,
                colourHex: ChartRedBandColourHex);

            AddBand(
                minimum: redTrigger,
                maximum: verticalMaximum,
                colourHex: ChartAboveRedBandColourHex);

            #endregion


            #region Draw Grid And Vertical Labels

            PenLineDefinition[] gridDefinitions =
            {
                new(verticalMaximum, verticalMaximum.ToString(CultureInfo.InvariantCulture)),
                new(redTrigger, redTrigger.ToString(CultureInfo.InvariantCulture)),
                new(amberTrigger, amberTrigger.ToString(CultureInfo.InvariantCulture)),
                new(greenTrigger, greenTrigger.ToString(CultureInfo.InvariantCulture)),
                new(0m, "0"),
                new(-greenTrigger, (-greenTrigger).ToString(CultureInfo.InvariantCulture)),
                new(-amberTrigger, (-amberTrigger).ToString(CultureInfo.InvariantCulture)),
                new(-redTrigger, (-redTrigger).ToString(CultureInfo.InvariantCulture)),
                new(verticalMinimum, verticalMinimum.ToString(CultureInfo.InvariantCulture))
            };

            foreach (PenLineDefinition gridDefinition
                in gridDefinitions)
            {
                double y =
                    ToY(
                        engineeringValue:
                            gridDefinition.Value);

                Line gridLine =
                    new()
                    {
                        X1 = plotLeft,
                        X2 = plotLeft + plotWidth,
                        Y1 = y,
                        Y2 = y,
                        Stroke =
                            gridDefinition.Value == 0m
                                ? Brushes.Black
                                : Brushes.LightGray,
                        StrokeThickness =
                            gridDefinition.Value == 0m
                                ? 1.8
                                : 0.7
                    };

                cnvChartPreview.Children.Add(
                    element: gridLine);

                Line verticalAxisTick =
                    new()
                    {
                        X1 = plotLeft - 5,
                        X2 = plotLeft,
                        Y1 = y,
                        Y2 = y,
                        Stroke = Brushes.Black,
                        StrokeThickness = 1.0
                    };

                cnvChartPreview.Children.Add(
                    element: verticalAxisTick);

                TextBlock label =
                    new()
                    {
                        Text =
                            gridDefinition.Label,

                        FontSize =
                            10,

                        Foreground =
                            Brushes.Black,

                        Width =
                            55,

                        TextAlignment =
                            TextAlignment.Right
                    };

                Canvas.SetLeft(
                    element: label,
                    length: 8);

                Canvas.SetTop(
                    element: label,
                    length: y - 8);

                cnvChartPreview.Children.Add(
                    element: label);
            }

            #endregion


            #region Draw Actual Series

            double previewDurationSeconds =
                Math.Max(
                    val1: 1.0,
                    val2: (_chartPreviewEndUtcExclusive - _chartPreviewStartUtc).TotalSeconds);

            foreach (ChartPreviewSeries previewSeries in _chartPreviewSeries)
            {
                System.Windows.Shapes.Polyline polyline =
                    new()
                    {
                        Stroke = previewSeries.Stroke,
                        StrokeThickness = previewSeries.LineWidth,
                        StrokeLineJoin = PenLineJoin.Round
                    };

                foreach (ChartPreviewPoint previewPoint in previewSeries.Points)
                {
                    double x =
                        plotLeft +
                        ((previewPoint.UtcTime - _chartPreviewStartUtc).TotalSeconds /
                         previewDurationSeconds * plotWidth);

                    decimal clippedValue =
                        Math.Min(
                            val1: verticalMaximum,
                            val2: Math.Max(
                                val1: verticalMinimum,
                                val2: Convert.ToDecimal(
                                    value: previewPoint.Value,
                                    provider: CultureInfo.InvariantCulture)));

                    polyline.Points.Add(
                        value:
                            new Point(
                                x: x,
                                y: ToY(
                                    engineeringValue: clippedValue)));
                }

                if (polyline.Points.Count > 0)
                {
                    cnvChartPreview.Children.Add(
                        element: polyline);
                }
            }

            #endregion


            #region Draw Border And Axes

            Line verticalAxis =
                new()
                {
                    X1 = plotLeft,
                    X2 = plotLeft,
                    Y1 = plotTop,
                    Y2 = plotBottom,
                    Stroke = Brushes.Black,
                    StrokeThickness = 1.0
                };

            cnvChartPreview.Children.Add(
                element: verticalAxis);

            Line horizontalAxis =
                new()
                {
                    X1 = plotLeft,
                    X2 = plotLeft + plotWidth,
                    Y1 = plotBottom,
                    Y2 = plotBottom,
                    Stroke = Brushes.Black,
                    StrokeThickness = 0.5
                };

            cnvChartPreview.Children.Add(
                element: horizontalAxis);

            #endregion


            #region Draw Chart Title

            string chartTitle =
                string.IsNullOrWhiteSpace(
                    value: txtChartTitleOverride.Text)
                    ? chartType.DisplayName
                    : txtChartTitleOverride.Text.Trim();

            TextBlock title =
                new()
                {
                    Text =
                        chartTitle,

                    FontWeight =
                        FontWeights.SemiBold,

                    FontSize =
                        15,

                    Width =
                        plotWidth,

                    TextAlignment =
                        TextAlignment.Center,

                    Foreground =
                        Brushes.Black
                };

            Canvas.SetLeft(
                element: title,
                length: plotLeft);

            Canvas.SetTop(
                element: title,
                length: 12);

            cnvChartPreview.Children.Add(
                element: title);

            #endregion


            #region Draw Vertical Axis Title

            string verticalAxisTitle =
                string.IsNullOrWhiteSpace(
                    value: txtChartYAxisTitle.Text)
                    ? chartType.VerticalAxisTitle
                    : txtChartYAxisTitle.Text.Trim();

            TextBlock verticalTitle =
                new()
                {
                    Text =
                        $"{verticalAxisTitle} ({chartType.Unit})",

                    FontSize =
                        11,

                    Foreground =
                        Brushes.Black,

                    RenderTransform =
                        new RotateTransform(
                            angle: -90),

                    RenderTransformOrigin =
                        new Point(
                            x: 0,
                            y: 0)
                };

            Canvas.SetLeft(
                element: verticalTitle,
                length: 5);

            Canvas.SetTop(
                element: verticalTitle,
                length: plotTop + (plotHeight * 0.72));

            cnvChartPreview.Children.Add(
                element: verticalTitle);

            #endregion


            #region Draw Horizontal Date Axis

            DateTime startDate =
                dpChartStartDate.SelectedDate!.Value.Date;

            DateTime endDate =
                dpChartEndDate.SelectedDate!.Value.Date;

            int previewDurationDays =
                Math.Max(
                    val1: 1,
                    val2: (endDate - startDate).Days);

            int tickIntervalDays =
                previewDurationDays <= 14
                    ? 1
                    : 7;

            int labelIntervalDays =
                previewDurationDays <= 14
                    ? 7
                    : 28;

            List<DateTime> tickDates =
                new();

            for (DateTime tickDate = startDate;
                 tickDate <= endDate;
                 tickDate = tickDate.AddDays(
                     value: tickIntervalDays))
            {
                tickDates.Add(
                    item: tickDate);
            }

            if (tickDates.Count == 0 ||
                tickDates[^1] != endDate)
            {
                tickDates.Add(
                    item: endDate);
            }

            foreach (DateTime tickDate in tickDates)
            {
                double fraction =
                    (tickDate - startDate).TotalDays /
                    previewDurationDays;

                double tickX =
                    plotLeft + (plotWidth * fraction);

                Line horizontalAxisTick =
                    new()
                    {
                        X1 = tickX,
                        X2 = tickX,
                        Y1 = plotBottom,
                        Y2 = plotBottom + 5,
                        Stroke = Brushes.Black,
                        StrokeThickness = 0.5
                    };

                cnvChartPreview.Children.Add(
                    element: horizontalAxisTick);
            }

            List<DateTime> labelDates =
                new();

            for (DateTime labelDate = startDate;
                 labelDate <= endDate;
                 labelDate = labelDate.AddDays(
                     value: labelIntervalDays))
            {
                labelDates.Add(
                    item: labelDate);
            }

            int minimumFinalLabelSpacingDays =
                labelIntervalDays / 2;

            if (labelDates.Count == 0)
            {
                labelDates.Add(
                    item: startDate);
            }

            if (labelDates[^1] != endDate &&
                (endDate - labelDates[^1]).Days >=
                minimumFinalLabelSpacingDays)
            {
                labelDates.Add(
                    item: endDate);
            }

            const double dateLabelWidth =
                82;

            foreach (DateTime labelDate in labelDates)
            {
                double fraction =
                    (labelDate - startDate).TotalDays /
                    previewDurationDays;

                double labelTickX =
                    plotLeft + (plotWidth * fraction);

                TextBlock label =
                    new()
                    {
                        Text =
                            labelDate.ToString(
                                format: "yyyy.MM.dd",
                                provider: CultureInfo.InvariantCulture),

                        FontSize =
                            10,

                        Width =
                            dateLabelWidth,

                        TextAlignment =
                            TextAlignment.Right,

                        RenderTransformOrigin =
                            new Point(
                                x: 1,
                                y: 0),

                        RenderTransform =
                            new RotateTransform(
                                angle: -45)
                    };

                Canvas.SetLeft(
                    element: label,
                    length: labelTickX - dateLabelWidth);

                Canvas.SetTop(
                    element: label,
                    length: plotBottom + 7);

                cnvChartPreview.Children.Add(
                    element: label);
            }

            TextBlock horizontalTitle =
                new()
                {
                    Text =
                        "Date",

                    FontSize =
                        11,

                    Width =
                        plotWidth,

                    TextAlignment =
                        TextAlignment.Center
                };

            Canvas.SetLeft(
                element: horizontalTitle,
                length: plotLeft);

            Canvas.SetTop(
                element: horizontalTitle,
                length: plotBottom + 72);

            cnvChartPreview.Children.Add(
                element: horizontalTitle);

            #endregion


            #region Draw Legend Placeholder

            if (chkChartShowLegend.IsChecked == true)
            {
                string legendText =
                    _chartSeries.Count > 0
                        ? string.Join(
                            separator: "    ",
                            values:
                                _chartSeries
                                    .OrderBy(
                                        keySelector:
                                            item => item.DisplayOrder)
                                    .Select(
                                        selector:
                                            item => item.DataElementDisplayName))
                        : "Legend";

                TextBlock legend =
                    new()
                    {
                        Text =
                            legendText,

                        FontSize =
                            10,

                        Foreground =
                            Brushes.Black,

                        Background =
                            Brushes.White,

                        Padding =
                            new Thickness(
                                uniformLength: 3)
                    };

                string legendPosition =
                    GetSelectedLegendPosition();

                if (string.Equals(
                    a: legendPosition,
                    b: "Below Title",
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    Canvas.SetLeft(
                        element: legend,
                        length: plotLeft + 8);

                    Canvas.SetTop(
                        element: legend,
                        length: 34);
                }
                else if (string.Equals(
                    a: legendPosition,
                    b: "Below Chart",
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    Canvas.SetLeft(
                        element: legend,
                        length: plotLeft + 8);

                    Canvas.SetTop(
                        element: legend,
                        length: plotBottom + 24);
                }
                else if (string.Equals(
                    a: legendPosition,
                    b: "Left of Chart",
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    Canvas.SetLeft(
                        element: legend,
                        length: 4);

                    Canvas.SetTop(
                        element: legend,
                        length: plotTop + 8);
                }
                else
                {
                    Canvas.SetLeft(
                        element: legend,
                        length: plotLeft + plotWidth - 120);

                    Canvas.SetTop(
                        element: legend,
                        length: plotTop + 8);
                }

                cnvChartPreview.Children.Add(
                    element: legend);
            }

            #endregion
        }


        private sealed record PenLineDefinition(
            decimal Value,
            string Label);


        private bool TryValidateChartConfiguration(
            out string validationMessage)
        {
            #region Validate Chart Header

            if (!_activeProjectId.HasValue)
            {
                validationMessage =
                    "Chart Management: Select an active project.";

                return false;
            }

            if (string.IsNullOrWhiteSpace(
                value: txtChartName.Text))
            {
                validationMessage =
                    "Definition: Enter a chart name.";

                return false;
            }

            if (cmbChartType.SelectedItem
                is not ChartTypeUiItem)
            {
                validationMessage =
                    "Definition: Select a chart type.";

                return false;
            }

            #endregion


            #region Validate Start Date Mode

            string startDateMode =
                GetSelectedChartStartDateMode();

            if (!string.Equals(
                    a: startDateMode,
                    b: ChartStartDateModeReportStart,
                    comparisonType: StringComparison.Ordinal) &&
                !string.Equals(
                    a: startDateMode,
                    b: ChartStartDateModeProjectStart,
                    comparisonType: StringComparison.Ordinal))
            {
                validationMessage =
                    "Time Basis: Select Report Start or Project Start.";

                return false;
            }

            #endregion


            #region Validate Trigger Levels

            if (!TryReadPositiveDecimal(
                text: txtChartGreenTrigger.Text,
                valueName: "Trigger Bands: Green trigger",
                value: out decimal greenTrigger,
                validationMessage: out validationMessage))
            {
                return false;
            }

            if (!TryReadPositiveDecimal(
                text: txtChartAmberTrigger.Text,
                valueName: "Trigger Bands: Amber trigger",
                value: out decimal amberTrigger,
                validationMessage: out validationMessage))
            {
                return false;
            }

            if (!TryReadPositiveDecimal(
                text: txtChartRedTrigger.Text,
                valueName: "Trigger Bands: Red trigger",
                value: out decimal redTrigger,
                validationMessage: out validationMessage))
            {
                return false;
            }

            if (amberTrigger <= greenTrigger)
            {
                validationMessage =
                    "Trigger Bands: Amber trigger must be greater than Green trigger.";

                return false;
            }

            if (redTrigger <= amberTrigger)
            {
                validationMessage =
                    "Trigger Bands: Red trigger must be greater than Amber trigger.";

                return false;
            }

            #endregion


            #region Validate Fixed Y Axis

            if (chkChartAutomaticYAxis.IsChecked != true)
            {
                if (!decimal.TryParse(
                    s: txtChartYAxisMinimum.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture,
                    result: out decimal yMinimum))
                {
                    validationMessage =
                        "Axes & Titles: Enter a valid numeric Vertical Axis minimum.";

                    return false;
                }

                if (!decimal.TryParse(
                    s: txtChartYAxisMaximum.Text,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture,
                    result: out decimal yMaximum))
                {
                    validationMessage =
                        "Axes & Titles: Enter a valid numeric Vertical Axis maximum.";

                    return false;
                }

                if (yMinimum >= yMaximum)
                {
                    validationMessage =
                        "Axes & Titles: Minimum must be less than Maximum.";

                    return false;
                }

                if (-redTrigger < yMinimum ||
                    redTrigger > yMaximum)
                {
                    validationMessage =
                        "Axes & Titles: Fixed Vertical Axis limits must include both Red trigger limits.";

                    return false;
                }
            }

            #endregion


            #region Validate Physical Output Dimensions

            if (!TryValidateChartAppearanceConfiguration(
                validationMessage: out validationMessage))
            {
                return false;
            }

            UpdateChartPixelDimensions();

            #endregion


            validationMessage =
                string.Empty;

            return true;
        }


        private static bool TryReadPositiveDecimal(
            string text,
            string valueName,
            out decimal value,
            out string validationMessage)
        {
            if (!decimal.TryParse(
                s: text,
                style: NumberStyles.Float,
                provider: CultureInfo.InvariantCulture,
                result: out value) ||
                value <= 0)
            {
                validationMessage =
                    $"{valueName} must be positive numeric.";

                return false;
            }

            validationMessage =
                string.Empty;

            return true;
        }


        private void btnChartRecentPeriod_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Apply Recent Period

            if (!int.TryParse(
                s: txtChartRecentDays.Text,
                style: NumberStyles.Integer,
                provider: CultureInfo.InvariantCulture,
                result: out int recentDays) ||
                recentDays <= 0)
            {
                txtChartStatus.Text =
                    "Recent days must be a positive whole number.";

                return;
            }

            DateTime endDate =
                dpChartEndDate.SelectedDate?.Date
                ?? DateTime.Today;

            dpChartEndDate.SelectedDate =
                endDate;

            dpChartStartDate.SelectedDate =
                endDate.AddDays(
                    value: -recentDays);

            txtChartStatus.Text =
                $"Chart Start updated to {recentDays} day(s) before End Date.";

            #endregion
        }


        private async void btnChartProjectStart_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Apply Active Project Start Date

            try
            {
                DateTime projectStartDate =
                    await GetActiveProjectStartDateAsync();

                dpChartStartDate.SelectedDate =
                    projectStartDate.Date;

                txtChartStatus.Text =
                    $"Report start date set to project start date: " +
                    $"{projectStartDate:yyyy-MM-dd}.";
            }
            catch (Exception ex)
            {
                txtChartStatus.Text =
                    $"Unable to read project start date: {ex.Message}";
            }

            #endregion
        }


        private void btnChartUseNow_Click(
            object sender,
            RoutedEventArgs e)
        {
            dpChartEndDate.SelectedDate =
                DateTime.Today;

            dpChartEndDate.Focus();

            txtChartStatus.Text =
                "Report end date set to Today.";
        }
        private async Task<DateTime> GetActiveProjectStartDateAsync()
        {
            #region Validate Active Project

            if (!_activeProjectId.HasValue)
            {
                throw new InvalidOperationException(
                    "Select an active project.");
            }

            #endregion


            #region Read Project Start Date

            const string sql = """
                SELECT [ProjectStartDate]
                FROM [dbo].[Project]
                WHERE
                    [Project_ID] = @Project_ID
                    AND [IsDeleted] = 0;
                """;

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            await using SqlCommand command =
                new(
                    cmdText: sql,
                    connection: databaseConnection);

            command.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    _activeProjectId.Value;

            object? value =
                await command.ExecuteScalarAsync();

            if (value is null ||
                value == DBNull.Value)
            {
                // ProjectStartDate is NOT NULL in the new schema, but retain a
                // defensive default for databases created before this revision.
                return DateTime.Today;
            }

            return Convert.ToDateTime(
                value: value,
                provider: CultureInfo.InvariantCulture)
                .Date;

            #endregion
        }


        private void chkChartAutomaticYAxis_Changed(
            object sender,
            RoutedEventArgs e)
        {
            UpdateChartYAxisControlState();
        }


        private void UpdateChartYAxisControlState()
        {
            if (!IsInitialized)
            {
                return;
            }

            bool fixedAxis =
                chkChartAutomaticYAxis.IsChecked != true;

            txtChartYAxisMinimum.IsEnabled =
                fixedAxis;

            txtChartYAxisMaximum.IsEnabled =
                fixedAxis;

            if (!fixedAxis)
            {
                ApplyDefaultVerticalAxisFromRedTrigger();
            }
        }


        private void btnChartAddSeries_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Add Series Placeholder

            if (cmbChartDataElement.SelectedItem
                is not string dataElement)
            {
                txtChartStatus.Text =
                    "Select a data element.";

                return;
            }

            if (cmbChartSeriesEntity.SelectedItem
                is not ChartEntityUiItem selectedEntity)
            {
                txtChartStatus.Text =
                    "Select an entity.";

                return;
            }

            string entityDisplayName =
                selectedEntity.DisplayName;

            if (_chartSeries.Any(
                predicate: series => series.EntityId == selectedEntity.EntityId))
            {
                txtChartStatus.Text =
                    $"Entity '{entityDisplayName}' already exists in this chart series.";

                return;
            }

            _chartSeries.Add(
                item:
                    new ChartSeriesUiItem
                    {
                        EntityId =
                            selectedEntity.EntityId,

                        DisplayOrder =
                            _chartSeries.Count + 1,

                        EntityDisplayName =
                            entityDisplayName,

                        DataElementDisplayName =
                            dataElement,

                        LegendText =
                            $"{entityDisplayName} - {dataElement}",

                        ColourHex =
                            string.Empty,

                        LineWidth =
                            ParsePositiveDoubleOrDefault(
                                text: txtChartDefaultLineWidth.Text,
                                defaultValue: 2.0),

                        MarkerSize =
                            ParseNonNegativeDoubleOrDefault(
                                text: txtChartDefaultMarkerSize.Text,
                                defaultValue: 4.0)
                    });

            if (cmbChartSeriesEntity.ItemsSource
                is IEnumerable<ChartEntityUiItem> availableEntities)
            {
                List<ChartEntityUiItem> remainingEntities =
                    availableEntities
                        .Where(
                            predicate: entity => entity.EntityId != selectedEntity.EntityId)
                        .OrderBy(
                            keySelector: entity => entity.DisplayName,
                            comparer: StringComparer.CurrentCultureIgnoreCase)
                        .ToList();

                cmbChartSeriesEntity.ItemsSource =
                    remainingEntities;

                cmbChartSeriesEntity.SelectedIndex =
                    remainingEntities.Count > 0
                        ? 0
                        : -1;
            }

            txtChartStatus.Text =
                $"Series for '{entityDisplayName}' added.";

            #endregion
        }


        private void btnChartRemoveSeries_Click(
            object sender,
            RoutedEventArgs e)
        {
            if (dgChartSeries.SelectedItem
                is not ChartSeriesUiItem selectedSeries)
            {
                return;
            }

            _chartSeries.Remove(
                item: selectedSeries);

            RenumberChartSeries();

            if (cmbChartSeriesEntity.ItemsSource
                is IEnumerable<ChartEntityUiItem> availableEntities)
            {
                List<ChartEntityUiItem> refreshedEntities =
                    availableEntities
                        .Append(
                            element:
                                new ChartEntityUiItem
                                {
                                    EntityId = selectedSeries.EntityId,
                                    DisplayName = selectedSeries.EntityDisplayName
                                })
                        .GroupBy(
                            keySelector: entity => entity.EntityId)
                        .Select(
                            selector: group => group.First())
                        .OrderBy(
                            keySelector: entity => entity.DisplayName,
                            comparer: StringComparer.CurrentCultureIgnoreCase)
                        .ToList();

                cmbChartSeriesEntity.ItemsSource =
                    refreshedEntities;

                cmbChartSeriesEntity.SelectedIndex =
                    refreshedEntities.Count > 0
                        ? 0
                        : -1;
            }
        }


        private void btnChartSeriesUp_Click(
            object sender,
            RoutedEventArgs e)
        {
            MoveSelectedChartSeries(
                direction: -1);
        }


        private void btnChartSeriesDown_Click(
            object sender,
            RoutedEventArgs e)
        {
            MoveSelectedChartSeries(
                direction: 1);
        }


        private void MoveSelectedChartSeries(
            int direction)
        {
            if (dgChartSeries.SelectedItem
                is not ChartSeriesUiItem selectedSeries)
            {
                return;
            }

            int oldIndex =
                _chartSeries.IndexOf(
                    item: selectedSeries);

            int newIndex =
                oldIndex + direction;

            if (newIndex < 0 ||
                newIndex >= _chartSeries.Count)
            {
                return;
            }

            _chartSeries.Move(
                oldIndex: oldIndex,
                newIndex: newIndex);

            RenumberChartSeries();

            dgChartSeries.SelectedItem =
                selectedSeries;
        }


        private void RenumberChartSeries()
        {
            for (int index = 0;
                 index < _chartSeries.Count;
                 index++)
            {
                _chartSeries[index].DisplayOrder =
                    index + 1;
            }

            dgChartSeries.Items.Refresh();
        }


        private static double ParsePositiveDoubleOrDefault(
            string text,
            double defaultValue)
        {
            return double.TryParse(
                s: text,
                style: NumberStyles.Float,
                provider: CultureInfo.InvariantCulture,
                result: out double value) &&
                value > 0
                    ? value
                    : defaultValue;
        }


        private static double ParseNonNegativeDoubleOrDefault(
            string text,
            double defaultValue)
        {
            return double.TryParse(
                s: text,
                style: NumberStyles.Float,
                provider: CultureInfo.InvariantCulture,
                result: out double value) &&
                value >= 0
                    ? value
                    : defaultValue;
        }

        #endregion


        #region Reference Coordinate Import

        #region CSV File Selection

        private async void btnSelectReferenceCsv_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Select CSV File

            OpenFileDialog openFileDialog =
                new()
                {
                    Title = "Select Point / Reference Coordinate CSV",

                    Filter =
                        "CSV files (*.csv)|*.csv|" +
                        "All files (*.*)|*.*",

                    DefaultExt =
                        ".csv",

                    CheckFileExists =
                        true,

                    Multiselect =
                        false
                };

            bool? fileSelected =
                openFileDialog.ShowDialog(
                    owner: this);

            if (fileSelected != true)
            {
                txtReferenceImportStatus.Text =
                    "CSV selection cancelled.";

                return;
            }

            #endregion


            #region Validate Active Project

            if (!_activeProjectId.HasValue)
            {
                txtReferenceImportStatus.Text =
                    "Import cannot continue because no active project is available.";

                MessageBox.Show(
                    owner: this,
                    messageBoxText:
                        "No active project is currently available.\n\n" +
                        "Select an active project using Manage Projects, " +
                        "then select the CSV again.",
                    caption: "Active Project Required",
                    button: MessageBoxButton.OK,
                    icon: MessageBoxImage.Warning);

                return;
            }

            if (string.IsNullOrWhiteSpace(
                value: _activeProjectName))
            {
                txtReferenceImportStatus.Text =
                    "Import cannot continue because the active project name is unavailable.";

                MessageBox.Show(
                    owner: this,
                    messageBoxText:
                        "The active project does not contain a valid project name.\n\n" +
                        "Select the project again using Manage Projects.",
                    caption: "Invalid Active Project",
                    button: MessageBoxButton.OK,
                    icon: MessageBoxImage.Warning);

                return;
            }

            #endregion


            #region Capture Import Project Context

            int importProjectId =
                _activeProjectId.Value;

            string importProjectName =
                _activeProjectName;

            _referenceImportProjectId =
                importProjectId;

            _referenceImportProjectName =
                importProjectName;

            #endregion


            #region Store Selected CSV

            _selectedReferenceCsvPath =
                openFileDialog.FileName;

            txtReferenceCsvPath.Text =
                _selectedReferenceCsvPath;

            #endregion


            #region Parse And Validate CSV

            await ParseAndValidateSelectedReferenceCsvAsync();

            #endregion
        }


        private async void cmbReferenceCsvColumnOrder_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Handle Initial XAML Loading

            if (string.IsNullOrWhiteSpace(
                value: _selectedReferenceCsvPath) ||
                !_referenceImportProjectId.HasValue)
            {
                return;
            }

            #endregion


            #region Reparse Selected CSV

            await ParseAndValidateSelectedReferenceCsvAsync();

            #endregion
        }


        private void btnClearReferenceImport_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Clear Import State

            ResetReferenceImportState(
                statusMessage: "No CSV selected.");

            #endregion
        }

        #endregion


        #region Import State Reset

        private void ResetReferenceImportState(
            string statusMessage)
        {
            #region Clear Captured Import Context

            _referenceImportProjectId =
                null;

            _referenceImportProjectName =
                string.Empty;

            _selectedReferenceCsvPath =
                string.Empty;

            #endregion


            #region Clear User Interface

            txtReferenceCsvPath.Text =
                string.Empty;

            _referenceImportItems.Clear();

            dgReferenceCoordinateImport.SelectedItem =
                null;

            btnCommitReferenceImport.IsEnabled =
                false;

            cmbReferenceCsvColumnOrder.IsEnabled =
                true;

            txtReferenceImportStatus.Text =
                statusMessage;

            #endregion
        }

        #endregion


        #region Import Preview Selection

        private void dgReferenceCoordinateImport_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Display Selected Row Validation

            if (dgReferenceCoordinateImport.SelectedItem
                is not ReferenceCoordinateImportItem selectedItem)
            {
                return;
            }

            txtReferenceImportStatus.Text =
                selectedItem.ValidationStatus;

            #endregion
        }

        #endregion


        #region Parse And Validate Selected CSV

        private async Task ParseAndValidateSelectedReferenceCsvAsync()
        {
            #region Validate Import Context

            if (!_referenceImportProjectId.HasValue)
            {
                throw new InvalidOperationException(
                    "The reference-coordinate import does not have a target Project_ID.");
            }

            if (string.IsNullOrWhiteSpace(_referenceImportProjectName))
            {
                throw new InvalidOperationException(
                    "The reference-coordinate import does not have a target project name.");
            }

            if (string.IsNullOrWhiteSpace(_selectedReferenceCsvPath))
            {
                throw new InvalidOperationException(
                    "No reference-coordinate CSV has been selected.");
            }

            if (!File.Exists(
                path: _selectedReferenceCsvPath))
            {
                throw new FileNotFoundException(
                    message: "The selected reference-coordinate CSV no longer exists.",
                    fileName: _selectedReferenceCsvPath);
            }

            #endregion


            #region Prepare User Interface

            btnSelectReferenceCsv.IsEnabled =
                false;

            btnClearReferenceImport.IsEnabled =
                false;

            cmbReferenceCsvColumnOrder.IsEnabled =
                false;

            btnCommitReferenceImport.IsEnabled =
                false;

            txtReferenceImportStatus.Text =
                $"Validating CSV for project '{_referenceImportProjectName}'...";

            #endregion


            try
            {
                #region Determine CSV Column Order

                ReferenceCoordinateCsvOrder csvOrder =
                    GetSelectedReferenceCsvOrder();

                #endregion


                #region Parse Source CSV

                List<ReferenceCoordinateImportItem> importItems =
                    ParseReferenceCoordinateCsv(
                        csvPath: _selectedReferenceCsvPath,
                        csvOrder: csvOrder);

                #endregion


                #region Load Existing Project Points

                List<ExistingReferencePointState> existingPoints =
                    await LoadExistingReferencePointStatesAsync(
                        projectId: _referenceImportProjectId.Value);

                #endregion


                #region Validate Import

                ValidateReferenceCoordinateImport(
                    importItems: importItems,
                    existingPoints: existingPoints);

                #endregion


                #region Populate Preview DataGrid

                _referenceImportItems.Clear();

                foreach (ReferenceCoordinateImportItem importItem in importItems)
                {
                    _referenceImportItems.Add(
                        item: importItem);
                }

                #endregion


                #region Determine Validation Result

                int invalidRowCount =
                    0;

                foreach (ReferenceCoordinateImportItem importItem in importItems)
                {
                    if (!importItem.IsValid)
                    {
                        invalidRowCount++;
                    }
                }

                if (importItems.Count == 0)
                {
                    btnCommitReferenceImport.IsEnabled =
                        false;

                    txtReferenceImportStatus.Text =
                        $"Project '{_referenceImportProjectName}': " +
                        "the selected CSV contains no records.";

                    return;
                }

                if (invalidRowCount > 0)
                {
                    btnCommitReferenceImport.IsEnabled =
                        false;

                    txtReferenceImportStatus.Text =
                        $"Project '{_referenceImportProjectName}': " +
                        $"{invalidRowCount} of {importItems.Count} row(s) contain errors. " +
                        "Import blocked. Correct the CSV and import the corrected file.";

                    return;
                }

                #endregion


                #region Report Successful Validation

                // The button may now be enabled because the preview contains a
                // completely valid import dataset.

                btnCommitReferenceImport.IsEnabled =
                    true;

                txtReferenceImportStatus.Text =
                    $"Project '{_referenceImportProjectName}': " +
                    $"{importItems.Count} row(s) validated successfully. " +
                    "No errors found.";

                #endregion
            }
            catch (Exception ex)
            {
                #region Handle Validation Failure

                _referenceImportItems.Clear();

                btnCommitReferenceImport.IsEnabled =
                    false;

                txtReferenceImportStatus.Text =
                    $"Unable to validate CSV: {ex.Message}";

                #endregion
            }
            finally
            {
                #region Restore User Interface

                btnSelectReferenceCsv.IsEnabled =
                    true;

                btnClearReferenceImport.IsEnabled =
                    true;

                cmbReferenceCsvColumnOrder.IsEnabled =
                    true;

                #endregion
            }
        }

        #endregion


        #region CSV Column Order

        private ReferenceCoordinateCsvOrder GetSelectedReferenceCsvOrder()
        {
            #region Read Selected Format

            return cmbReferenceCsvColumnOrder.SelectedIndex switch
            {
                0 => ReferenceCoordinateCsvOrder.PointName_E_N_Ht,

                1 => ReferenceCoordinateCsvOrder.PointName_N_E_Ht,

                _ => throw new InvalidOperationException(
                    "Select a valid CSV column order.")
            };

            #endregion
        }

        #endregion


        #region CSV Parsing

        private static List<ReferenceCoordinateImportItem>
            ParseReferenceCoordinateCsv(
                string csvPath,
                ReferenceCoordinateCsvOrder csvOrder)
        {
            #region Validate CSV Path

            if (string.IsNullOrWhiteSpace(csvPath))
            {
                throw new ArgumentException(
                    message: "CSV path cannot be empty.",
                    paramName: nameof(csvPath));
            }

            #endregion


            #region Read CSV Lines

            string[] sourceLines =
                File.ReadAllLines(
                    path: csvPath);

            List<ReferenceCoordinateImportItem> importItems =
                new();

            #endregion


            #region Parse CSV Rows

            for (int lineIndex = 0;
                 lineIndex < sourceLines.Length;
                 lineIndex++)
            {
                int sourceRow =
                    lineIndex + 1;

                string rawLine =
                    sourceLines[lineIndex];

                ReferenceCoordinateImportItem importItem =
                    new()
                    {
                        SourceRow = sourceRow,
                        RawLine = rawLine,
                        IsValid = true,
                        ValidationStatus = "Valid"
                    };


                #region Parse CSV Fields

                if (!TryParseCsvLine(
                    line: rawLine,
                    fields: out List<string> fields,
                    errorMessage: out string csvError))
                {
                    AppendValidationError(
                        importItem: importItem,
                        errorMessage: csvError);

                    importItems.Add(
                        item: importItem);

                    continue;
                }

                #endregion


                #region Validate Column Count

                if (fields.Count != 4 &&
                    fields.Count != 5)
                {
                    AppendValidationError(
                        importItem: importItem,
                        errorMessage:
                            $"Expected 4 or 5 columns; found {fields.Count}.");

                    importItems.Add(
                        item: importItem);

                    continue;
                }

                #endregion


                #region Extract Point Names

                importItem.PointName =
                    fields[0].Trim();

                string replacementName =
                    fields.Count == 5
                        ? fields[4].Trim()
                        : string.Empty;

                importItem.ReplacementName =
                    string.IsNullOrWhiteSpace(replacementName)
                        ? importItem.PointName
                        : replacementName;

                #endregion


                #region Extract Coordinate Text

                switch (csvOrder)
                {
                    case ReferenceCoordinateCsvOrder.PointName_E_N_Ht:

                        importItem.EastingText =
                            fields[1].Trim();

                        importItem.NorthingText =
                            fields[2].Trim();

                        break;


                    case ReferenceCoordinateCsvOrder.PointName_N_E_Ht:

                        importItem.NorthingText =
                            fields[1].Trim();

                        importItem.EastingText =
                            fields[2].Trim();

                        break;


                    default:

                        throw new InvalidOperationException(
                            "Unsupported reference-coordinate CSV column order.");
                }

                importItem.HeightText =
                    fields[3].Trim();

                #endregion


                #region Validate Point Name

                if (string.IsNullOrWhiteSpace(importItem.PointName))
                {
                    AppendValidationError(
                        importItem: importItem,
                        errorMessage: "Point Name is blank.");
                }
                else if (importItem.PointName.Length > 50)
                {
                    AppendValidationError(
                        importItem: importItem,
                        errorMessage:
                            "Point Name exceeds the database limit of 50 characters.");
                }

                #endregion


                #region Validate Replacement Name

                if (string.IsNullOrWhiteSpace(importItem.ReplacementName))
                {
                    AppendValidationError(
                        importItem: importItem,
                        errorMessage: "Replacement Name is blank.");
                }
                else if (importItem.ReplacementName.Length > 50)
                {
                    AppendValidationError(
                        importItem: importItem,
                        errorMessage:
                            "Replacement Name exceeds the database limit of 50 characters.");
                }

                #endregion


                #region Parse Easting

                if (decimal.TryParse(
                    s: importItem.EastingText,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture,
                    result: out decimal easting))
                {
                    importItem.Easting =
                        easting;
                }
                else
                {
                    AppendValidationError(
                        importItem: importItem,
                        errorMessage:
                            $"Invalid Easting '{importItem.EastingText}'.");
                }

                #endregion


                #region Parse Northing

                if (decimal.TryParse(
                    s: importItem.NorthingText,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture,
                    result: out decimal northing))
                {
                    importItem.Northing =
                        northing;
                }
                else
                {
                    AppendValidationError(
                        importItem: importItem,
                        errorMessage:
                            $"Invalid Northing '{importItem.NorthingText}'.");
                }

                #endregion


                #region Parse Height

                if (decimal.TryParse(
                    s: importItem.HeightText,
                    style: NumberStyles.Float,
                    provider: CultureInfo.InvariantCulture,
                    result: out decimal height))
                {
                    importItem.Height =
                        height;
                }
                else
                {
                    AppendValidationError(
                        importItem: importItem,
                        errorMessage:
                            $"Invalid Height '{importItem.HeightText}'.");
                }

                #endregion


                #region Add Parsed Row

                importItems.Add(
                    item: importItem);

                #endregion
            }

            #endregion


            #region Return Parsed CSV

            return importItems;

            #endregion
        }


        private static bool TryParseCsvLine(
            string line,
            out List<string> fields,
            out string errorMessage)
        {
            #region Initialise Parser

            fields =
                new List<string>();

            errorMessage =
                string.Empty;

            StringBuilder currentField =
                new();

            bool insideQuotes =
                false;

            #endregion


            #region Parse Characters

            for (int characterIndex = 0;
                 characterIndex < line.Length;
                 characterIndex++)
            {
                char currentCharacter =
                    line[characterIndex];

                if (currentCharacter == '"')
                {
                    if (insideQuotes &&
                        characterIndex + 1 < line.Length &&
                        line[characterIndex + 1] == '"')
                    {
                        currentField.Append(
                            value: '"');

                        characterIndex++;

                        continue;
                    }

                    insideQuotes =
                        !insideQuotes;

                    continue;
                }

                if (currentCharacter == ',' &&
                    !insideQuotes)
                {
                    fields.Add(
                        item: currentField.ToString());

                    currentField.Clear();

                    continue;
                }

                currentField.Append(
                    value: currentCharacter);
            }

            #endregion


            #region Validate Quotation

            if (insideQuotes)
            {
                errorMessage =
                    "Malformed CSV row: unmatched quotation mark.";

                return false;
            }

            #endregion


            #region Add Final Field

            fields.Add(
                item: currentField.ToString());

            return true;

            #endregion
        }

        #endregion


        #region Load Existing Project Reference Points

        private async Task<List<ExistingReferencePointState>>
            LoadExistingReferencePointStatesAsync(
                int projectId)
        {
            #region Validate Project ID

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(projectId),
                    message: "Project_ID must be greater than zero.");
            }

            #endregion


            #region Open Database Connection

            string databaseConnectionString =
                GetTrackGeometryConnectionString();

            await using SqlConnection databaseConnection =
                new(
                    connectionString: databaseConnectionString);

            await databaseConnection.OpenAsync();

            #endregion


            #region Load Reference Point State

            return await LoadExistingReferencePointStatesAsync(
                projectId: projectId,
                databaseConnection: databaseConnection,
                transaction: null,
                applyCommitLocks: false);

            #endregion
        }


        private static async Task<List<ExistingReferencePointState>>
            LoadExistingReferencePointStatesAsync(
                int projectId,
                SqlConnection databaseConnection,
                SqlTransaction? transaction,
                bool applyCommitLocks)
        {
            #region Define Existing Point Query

            string pointLockHint =
                applyCommitLocks
                    ? " WITH (UPDLOCK, HOLDLOCK)"
                    : string.Empty;

            string referenceLockHint =
                applyCommitLocks
                    ? " WITH (UPDLOCK, HOLDLOCK)"
                    : string.Empty;

            string existingPointsSql =
                $"""
                SELECT
                    PN.[PointName_ID],
                    PN.[Project_ID],
                    PN.[PointName],
                    PN.[ReplacementName],
                    PN.[IsDeleted],
                    CR.[Eref],
                    CR.[Nref],
                    CR.[Href]
                FROM [dbo].[PointName] AS PN{pointLockHint}
                LEFT JOIN [dbo].[CoordinatesReference] AS CR{referenceLockHint}
                    ON CR.[PointName_ID] = PN.[PointName_ID]
                    AND CR.[IsDeleted] = 0
                WHERE
                    PN.[Project_ID] = @Project_ID;
                """;

            #endregion


            #region Load Existing Point State

            List<ExistingReferencePointState> existingPoints =
                new();

            await using SqlCommand existingPointsCommand =
                transaction is null
                    ? new SqlCommand(
                        cmdText: existingPointsSql,
                        connection: databaseConnection)
                    : new SqlCommand(
                        cmdText: existingPointsSql,
                        connection: databaseConnection,
                        transaction: transaction);

            existingPointsCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            await using SqlDataReader reader =
                await existingPointsCommand.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                bool hasReference =
                    !reader.IsDBNull(
                        i: 5);

                ExistingReferencePointState existingPoint =
                    new()
                    {
                        PointName_ID =
                            reader.GetInt32(
                                i: 0),

                        Project_ID =
                            reader.GetInt32(
                                i: 1),

                        PointName =
                            reader.GetString(
                                i: 2),

                        ReplacementName =
                            reader.GetString(
                                i: 3),

                        IsDeleted =
                            reader.GetBoolean(
                                i: 4),

                        HasReference =
                            hasReference,

                        Easting =
                            hasReference
                                ? reader.GetDecimal(
                                    i: 5)
                                : null,

                        Northing =
                            hasReference
                                ? reader.GetDecimal(
                                    i: 6)
                                : null,

                        Height =
                            hasReference
                                ? reader.GetDecimal(
                                    i: 7)
                                : null
                    };

                existingPoints.Add(
                    item: existingPoint);
            }

            return existingPoints;

            #endregion
        }


        private sealed class ExistingReferencePointState
        {
            public int PointName_ID { get; init; }

            public int Project_ID { get; init; }

            public string PointName { get; init; } =
                string.Empty;

            public string ReplacementName { get; init; } =
                string.Empty;

            public bool IsDeleted { get; init; }

            public bool HasReference { get; init; }

            public decimal? Easting { get; init; }

            public decimal? Northing { get; init; }

            public decimal? Height { get; init; }
        }

        #endregion


        #region Complete Import Validation

        private static void ValidateReferenceCoordinateImport(
            List<ReferenceCoordinateImportItem> importItems,
            List<ExistingReferencePointState> existingPoints)
        {
            #region Validate Arguments

            ArgumentNullException.ThrowIfNull(
                argument: importItems);

            ArgumentNullException.ThrowIfNull(
                argument: existingPoints);

            #endregion


            #region Validate Combined Names Within CSV

            Dictionary<string,
                (ReferenceCoordinateImportItem Item, string NameType)>
                incomingNames =
                    new(
                        comparer: StringComparer.OrdinalIgnoreCase);

            foreach (ReferenceCoordinateImportItem importItem in importItems)
            {
                RegisterIncomingName(
                    name: importItem.PointName,
                    nameType: "Point Name",
                    importItem: importItem,
                    incomingNames: incomingNames);

                if (!string.Equals(
                    a: importItem.PointName,
                    b: importItem.ReplacementName,
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    RegisterIncomingName(
                        name: importItem.ReplacementName,
                        nameType: "Replacement Name",
                        importItem: importItem,
                        incomingNames: incomingNames);
                }
            }

            #endregion


            #region Build Existing Point Namespace

            Dictionary<string, List<ExistingReferencePointState>> existingNames =
                BuildExistingReferencePointNamespace(
                    existingPoints: existingPoints);

            #endregion


            #region Validate Incoming Names Against Database

            foreach (ReferenceCoordinateImportItem importItem in importItems)
            {
                ValidateIncomingPointAgainstExistingDatabase(
                    importItem: importItem,
                    existingNames: existingNames);
            }

            #endregion


            #region Validate Horizontal Coordinate Duplicates

            ValidateHorizontalCoordinateDuplicates(
                importItems: importItems,
                existingPoints: existingPoints);

            #endregion


            #region Finalise Validation Status

            foreach (ReferenceCoordinateImportItem importItem in importItems)
            {
                if (importItem.IsValid)
                {
                    importItem.ValidationStatus =
                        "Valid";
                }
            }

            #endregion
        }

        #endregion


        #region Name Validation Helpers

        private static void RegisterIncomingName(
            string name,
            string nameType,
            ReferenceCoordinateImportItem importItem,
            Dictionary<string,
                (ReferenceCoordinateImportItem Item, string NameType)> incomingNames)
        {
            #region Ignore Invalid Blank Names

            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            #endregion


            #region Check Existing Incoming Name

            if (incomingNames.TryGetValue(
                key: name,
                value: out
                    (ReferenceCoordinateImportItem Item, string NameType)
                    existingName))
            {
                AppendValidationError(
                    importItem: importItem,
                    errorMessage:
                        $"{nameType} '{name}' duplicates CSV row " +
                        $"{existingName.Item.SourceRow}.");

                AppendValidationError(
                    importItem: existingName.Item,
                    errorMessage:
                        $"'{name}' duplicates CSV row {importItem.SourceRow}.");

                return;
            }

            #endregion


            #region Register Name

            incomingNames.Add(
                key: name,
                value: (importItem, nameType));

            #endregion
        }


        private static Dictionary<string, List<ExistingReferencePointState>>
            BuildExistingReferencePointNamespace(
                List<ExistingReferencePointState> existingPoints)
        {
            #region Build Existing Namespace

            Dictionary<string, List<ExistingReferencePointState>> existingNames =
                new(
                    comparer: StringComparer.OrdinalIgnoreCase);

            foreach (ExistingReferencePointState existingPoint in existingPoints)
            {
                RegisterExistingReferencePointName(
                    name: existingPoint.PointName,
                    existingPoint: existingPoint,
                    existingNames: existingNames);

                if (!string.Equals(
                    a: existingPoint.PointName,
                    b: existingPoint.ReplacementName,
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    RegisterExistingReferencePointName(
                        name: existingPoint.ReplacementName,
                        existingPoint: existingPoint,
                        existingNames: existingNames);
                }
            }

            return existingNames;

            #endregion
        }


        private static void RegisterExistingReferencePointName(
            string name,
            ExistingReferencePointState existingPoint,
            Dictionary<string, List<ExistingReferencePointState>> existingNames)
        {
            #region Register Existing Name

            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            if (!existingNames.TryGetValue(
                key: name,
                value: out List<ExistingReferencePointState>? matches))
            {
                matches =
                    new();

                existingNames.Add(
                    key: name,
                    value: matches);
            }

            foreach (ExistingReferencePointState match in matches)
            {
                if (match.PointName_ID == existingPoint.PointName_ID)
                {
                    return;
                }
            }

            matches.Add(
                item: existingPoint);

            #endregion
        }


        private static void ValidateIncomingPointAgainstExistingDatabase(
            ReferenceCoordinateImportItem importItem,
            Dictionary<string, List<ExistingReferencePointState>> existingNames)
        {
            #region Resolve Existing Database Point

            Dictionary<int, ExistingReferencePointState> matchesByPointId =
                new();

            AddExistingReferenceMatches(
                lookupName: importItem.PointName,
                existingNames: existingNames,
                matchesByPointId: matchesByPointId);

            AddExistingReferenceMatches(
                lookupName: importItem.ReplacementName,
                existingNames: existingNames,
                matchesByPointId: matchesByPointId);

            if (matchesByPointId.Count > 1)
            {
                AppendValidationError(
                    importItem: importItem,
                    errorMessage: "Name collision in database.");

                return;
            }

            ExistingReferencePointState? existingPoint =
                null;

            foreach (ExistingReferencePointState match
                in matchesByPointId.Values)
            {
                existingPoint =
                    match;

                break;
            }

            if (existingPoint is null)
            {
                return;
            }

            #endregion


            #region Validate Existing Mapping

            if (existingPoint.IsDeleted)
            {
                AppendValidationError(
                    importItem: importItem,
                    errorMessage:
                        $"'{importItem.PointName}': Point is deleted.");

                return;
            }

            bool pointNameMatches =
                string.Equals(
                    a: existingPoint.PointName,
                    b: importItem.PointName,
                    comparisonType: StringComparison.OrdinalIgnoreCase);

            bool replacementNameMatches =
                string.Equals(
                    a: existingPoint.ReplacementName,
                    b: importItem.ReplacementName,
                    comparisonType: StringComparison.OrdinalIgnoreCase);

            if (!pointNameMatches ||
                !replacementNameMatches)
            {
                AppendValidationError(
                    importItem: importItem,
                    errorMessage:
                        $"'{importItem.PointName}': Name mapping differs.");

                return;
            }

            if (existingPoint.HasReference)
            {
                AppendValidationError(
                    importItem: importItem,
                    errorMessage:
                        $"'{importItem.PointName}': Reference coordinates already exist.");
            }

            #endregion
        }


        private static void AddExistingReferenceMatches(
            string lookupName,
            Dictionary<string, List<ExistingReferencePointState>> existingNames,
            Dictionary<int, ExistingReferencePointState> matchesByPointId)
        {
            #region Add Existing Matches

            if (string.IsNullOrWhiteSpace(lookupName) ||
                !existingNames.TryGetValue(
                    key: lookupName,
                    value: out List<ExistingReferencePointState>? matches))
            {
                return;
            }

            foreach (ExistingReferencePointState match in matches)
            {
                matchesByPointId.TryAdd(
                    key: match.PointName_ID,
                    value: match);
            }

            #endregion
        }
        #endregion


        #region Horizontal Coordinate Duplicate Validation

        private static void ValidateHorizontalCoordinateDuplicates(
            List<ReferenceCoordinateImportItem> importItems,
            List<ExistingReferencePointState> existingPoints)
        {
            #region Build Existing Coordinate Buckets

            Dictionary<(long EastingBucket, long NorthingBucket),
                List<ExistingReferencePointState>>
                existingCoordinateBuckets =
                    new();

            foreach (ExistingReferencePointState existingPoint in existingPoints)
            {
                if (!existingPoint.HasReference ||
                    !existingPoint.Easting.HasValue ||
                    !existingPoint.Northing.HasValue)
                {
                    continue;
                }

                (long EastingBucket, long NorthingBucket) bucket =
                    GetReferenceCoordinateBucket(
                        easting: existingPoint.Easting.Value,
                        northing: existingPoint.Northing.Value);

                if (!existingCoordinateBuckets.TryGetValue(
                    key: bucket,
                    value: out List<ExistingReferencePointState>? bucketPoints))
                {
                    bucketPoints =
                        new List<ExistingReferencePointState>();

                    existingCoordinateBuckets.Add(
                        key: bucket,
                        value: bucketPoints);
                }

                bucketPoints.Add(
                    item: existingPoint);
            }

            #endregion


            #region Prepare Incoming Coordinate Buckets

            Dictionary<(long EastingBucket, long NorthingBucket),
                List<ReferenceCoordinateImportItem>>
                incomingCoordinateBuckets =
                    new();

            #endregion


            #region Validate Each Incoming Coordinate

            foreach (ReferenceCoordinateImportItem importItem in importItems)
            {
                if (!importItem.Easting.HasValue ||
                    !importItem.Northing.HasValue)
                {
                    continue;
                }

                decimal easting =
                    importItem.Easting.Value;

                decimal northing =
                    importItem.Northing.Value;

                (long EastingBucket, long NorthingBucket) sourceBucket =
                    GetReferenceCoordinateBucket(
                        easting: easting,
                        northing: northing);


                #region Check Against Existing Database Points

                for (long eastOffset = -1;
                     eastOffset <= 1;
                     eastOffset++)
                {
                    for (long northOffset = -1;
                         northOffset <= 1;
                         northOffset++)
                    {
                        (long EastingBucket, long NorthingBucket) neighbourBucket =
                            (
                                sourceBucket.EastingBucket + eastOffset,
                                sourceBucket.NorthingBucket + northOffset
                            );

                        if (!existingCoordinateBuckets.TryGetValue(
                            key: neighbourBucket,
                            value: out
                                List<ExistingReferencePointState>? existingBucketPoints))
                        {
                            continue;
                        }

                        foreach (ExistingReferencePointState existingPoint
                            in existingBucketPoints)
                        {
                            decimal distanceSquared =
                                GetHorizontalDistanceSquared(
                                    easting1: easting,
                                    northing1: northing,
                                    easting2: existingPoint.Easting.Value,
                                    northing2: existingPoint.Northing.Value);

                            if (distanceSquared <
                                ReferenceDuplicateCoordinateToleranceSquared)
                            {
                                double distance =
                                    Math.Sqrt(
                                        d: (double)distanceSquared);

                                AppendValidationError(
                                    importItem: importItem,
                                    errorMessage:
                                        $"Coordinates are {distance.ToString(
                                            format: "0.000",
                                            provider: CultureInfo.InvariantCulture)} m " +
                                        $"from existing database point " +
                                        $"'{existingPoint.PointName}'.");
                            }
                        }
                    }
                }

                #endregion


                #region Check Against Earlier CSV Points

                for (long eastOffset = -1;
                     eastOffset <= 1;
                     eastOffset++)
                {
                    for (long northOffset = -1;
                         northOffset <= 1;
                         northOffset++)
                    {
                        (long EastingBucket, long NorthingBucket) neighbourBucket =
                            (
                                sourceBucket.EastingBucket + eastOffset,
                                sourceBucket.NorthingBucket + northOffset
                            );

                        if (!incomingCoordinateBuckets.TryGetValue(
                            key: neighbourBucket,
                            value: out
                                List<ReferenceCoordinateImportItem>?
                                incomingBucketPoints))
                        {
                            continue;
                        }

                        foreach (ReferenceCoordinateImportItem previousItem
                            in incomingBucketPoints)
                        {
                            if (!previousItem.Easting.HasValue ||
                                !previousItem.Northing.HasValue)
                            {
                                continue;
                            }

                            decimal distanceSquared =
                                GetHorizontalDistanceSquared(
                                    easting1: easting,
                                    northing1: northing,
                                    easting2: previousItem.Easting.Value,
                                    northing2: previousItem.Northing.Value);

                            if (distanceSquared <
                                ReferenceDuplicateCoordinateToleranceSquared)
                            {
                                double distance =
                                    Math.Sqrt(
                                        d: (double)distanceSquared);

                                string formattedDistance =
                                    distance.ToString(
                                        format: "0.000",
                                        provider: CultureInfo.InvariantCulture);

                                AppendValidationError(
                                    importItem: importItem,
                                    errorMessage:
                                        $"Coordinates are {formattedDistance} m " +
                                        $"from CSV row {previousItem.SourceRow} " +
                                        $"('{previousItem.PointName}').");

                                AppendValidationError(
                                    importItem: previousItem,
                                    errorMessage:
                                        $"Coordinates are {formattedDistance} m " +
                                        $"from CSV row {importItem.SourceRow} " +
                                        $"('{importItem.PointName}').");
                            }
                        }
                    }
                }

                #endregion


                #region Add Incoming Point To Spatial Bucket

                if (!incomingCoordinateBuckets.TryGetValue(
                    key: sourceBucket,
                    value: out
                        List<ReferenceCoordinateImportItem>? sourceBucketPoints))
                {
                    sourceBucketPoints =
                        new List<ReferenceCoordinateImportItem>();

                    incomingCoordinateBuckets.Add(
                        key: sourceBucket,
                        value: sourceBucketPoints);
                }

                sourceBucketPoints.Add(
                    item: importItem);

                #endregion
            }

            #endregion
        }


        private static (
            long EastingBucket,
            long NorthingBucket)
            GetReferenceCoordinateBucket(
                decimal easting,
                decimal northing)
        {
            #region Calculate Spatial Bucket

            long eastingBucket =
                (long)decimal.Floor(
                    d: easting /
                       ReferenceDuplicateCoordinateTolerance);

            long northingBucket =
                (long)decimal.Floor(
                    d: northing /
                       ReferenceDuplicateCoordinateTolerance);

            return
                (
                    EastingBucket: eastingBucket,
                    NorthingBucket: northingBucket
                );

            #endregion
        }


        private static decimal GetHorizontalDistanceSquared(
            decimal easting1,
            decimal northing1,
            decimal easting2,
            decimal northing2)
        {
            #region Calculate Horizontal Distance Squared

            decimal deltaEasting =
                easting2 - easting1;

            decimal deltaNorthing =
                northing2 - northing1;

            return
                (deltaEasting * deltaEasting) +
                (deltaNorthing * deltaNorthing);

            #endregion
        }

        #endregion


        #region Validation Status Helper

        private static void AppendValidationError(
            ReferenceCoordinateImportItem importItem,
            string errorMessage)
        {
            #region Validate Parameters

            ArgumentNullException.ThrowIfNull(
                argument: importItem);

            if (string.IsNullOrWhiteSpace(errorMessage))
            {
                return;
            }

            #endregion


            #region Append Error

            importItem.IsValid =
                false;

            if (string.IsNullOrWhiteSpace(importItem.ValidationStatus) ||
                string.Equals(
                    a: importItem.ValidationStatus,
                    b: "Valid",
                    comparisonType: StringComparison.Ordinal))
            {
                importItem.ValidationStatus =
                    errorMessage;
            }
            else
            {
                importItem.ValidationStatus =
                    $"{importItem.ValidationStatus}; {errorMessage}";
            }

            #endregion
        }

        #endregion


        #region Commit Reference Coordinate Import

        private async void btnCommitReferenceImport_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Commit Context

            if (!_referenceImportProjectId.HasValue)
            {
                btnCommitReferenceImport.IsEnabled =
                    false;

                txtReferenceImportStatus.Text =
                    "Import blocked: no target project is associated with the preview.";

                return;
            }

            if (!_activeProjectId.HasValue ||
                _activeProjectId.Value != _referenceImportProjectId.Value)
            {
                btnCommitReferenceImport.IsEnabled =
                    false;

                txtReferenceImportStatus.Text =
                    "Import blocked: the active project has changed. Select the CSV again.";

                return;
            }

            if (_referenceImportItems.Count == 0)
            {
                btnCommitReferenceImport.IsEnabled =
                    false;

                txtReferenceImportStatus.Text =
                    "Import blocked: there are no reference points to import.";

                return;
            }

            foreach (ReferenceCoordinateImportItem importItem
                in _referenceImportItems)
            {
                if (!importItem.IsValid)
                {
                    btnCommitReferenceImport.IsEnabled =
                        false;

                    txtReferenceImportStatus.Text =
                        "Import blocked: the preview contains validation errors.";

                    return;
                }
            }

            #endregion


            #region Confirm Import

            txtReferenceImportStatus.Text =
                $"Ready to import {_referenceImportItems.Count} point(s) " +
                $"into project '{_referenceImportProjectName}'.";

            string confirmationMessage =
                $"Project: {_referenceImportProjectName}\n\n" +
                $"Points to import: {_referenceImportItems.Count}\n" +
                "Validation errors: 0\n\n" +
                "Commit this import?";

            MessageBoxResult confirmationResult =
                MessageBox.Show(
                    owner: this,
                    messageBoxText: confirmationMessage,
                    caption: "Confirm Reference Coordinate Import",
                    button: MessageBoxButton.YesNo,
                    icon: MessageBoxImage.Question);

            if (confirmationResult != MessageBoxResult.Yes)
            {
                txtReferenceImportStatus.Text =
                    "Import cancelled. No data was written to the database.";

                return;
            }

            #endregion


            #region Prepare Commit

            btnCommitReferenceImport.IsEnabled =
                false;

            btnSelectReferenceCsv.IsEnabled =
                false;

            btnClearReferenceImport.IsEnabled =
                false;

            cmbReferenceCsvColumnOrder.IsEnabled =
                false;

            txtReferenceImportStatus.Text =
                $"Revalidating and importing {_referenceImportItems.Count} point(s) " +
                $"into project '{_referenceImportProjectName}'...";

            #endregion


            try
            {
                #region Execute Atomic Import

                int importedPointCount =
                    await CommitReferenceCoordinateImportAsync(
                        projectId: _referenceImportProjectId.Value,
                        projectName: _referenceImportProjectName);

                #endregion


                #region Report Successful Import

                btnCommitReferenceImport.IsEnabled =
                    false;

                cmbReferenceCsvColumnOrder.IsEnabled =
                    false;

                txtReferenceImportStatus.Text =
                    $"{importedPointCount} point(s) imported successfully " +
                    $"into project '{_referenceImportProjectName}'.";

                #endregion
            }
            catch (Exception ex)
            {
                #region Report Failed Import

                bool previewStillValid =
                    true;

                foreach (ReferenceCoordinateImportItem importItem
                    in _referenceImportItems)
                {
                    if (!importItem.IsValid)
                    {
                        previewStillValid =
                            false;

                        break;
                    }
                }

                btnCommitReferenceImport.IsEnabled =
                    previewStillValid;

                dgReferenceCoordinateImport.Items.Refresh();

                txtReferenceImportStatus.Text =
                    $"Import failed. No data committed. {ex.Message}";

                #endregion
            }
            finally
            {
                #region Restore Import Controls

                btnSelectReferenceCsv.IsEnabled =
                    true;

                btnClearReferenceImport.IsEnabled =
                    true;

                #endregion
            }
        }

        #endregion


        #region Atomic Reference Coordinate Database Import

        private async Task<int> CommitReferenceCoordinateImportAsync(
            int projectId,
            string projectName)
        {
            #region Validate Parameters

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(projectId),
                    message: "Project_ID must be greater than zero.");
            }

            if (string.IsNullOrWhiteSpace(projectName))
            {
                throw new ArgumentException(
                    message: "Project name cannot be blank.",
                    paramName: nameof(projectName));
            }

            #endregion


            #region Capture Preview Dataset

            List<ReferenceCoordinateImportItem> importItems =
                new(
                    collection: _referenceImportItems);

            if (importItems.Count == 0)
            {
                throw new InvalidOperationException(
                    "The reference-coordinate preview contains no records.");
            }

            #endregion


            #region Resolve Database Connection String

            string databaseConnectionString =
                GetTrackGeometryConnectionString();

            #endregion


            #region Open Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionString);

            await databaseConnection.OpenAsync();

            #endregion


            #region Begin Serializable Transaction

            using SqlTransaction transaction =
                databaseConnection.BeginTransaction(
                    iso:
                        System.Data.IsolationLevel.Serializable);

            bool transactionCommitted =
                false;

            #endregion


            try
            {
                #region Acquire Import Transaction Lock

                await AcquireReferenceImportTransactionLockAsync(
                    projectId: projectId,
                    databaseConnection: databaseConnection,
                    transaction: transaction);

                #endregion


                #region Revalidate Target Project

                await ValidateImportTargetProjectAsync(
                    projectId: projectId,
                    expectedProjectName: projectName,
                    databaseConnection: databaseConnection,
                    transaction: transaction);

                #endregion


                #region Reset Preview Validation Status

                foreach (ReferenceCoordinateImportItem importItem
                    in importItems)
                {
                    importItem.IsValid =
                        true;

                    importItem.ValidationStatus =
                        "Valid";
                }

                #endregion


                #region Reload Current Database Reference Points

                List<ExistingReferencePointState> existingPoints =
                    await LoadExistingReferencePointStatesAsync(
                        projectId: projectId,
                        databaseConnection: databaseConnection,
                        transaction: transaction,
                        applyCommitLocks: true);

                #endregion


                #region Revalidate Import Against Current Database

                ValidateReferenceCoordinateImport(
                    importItems: importItems,
                    existingPoints: existingPoints);

                int invalidRowCount =
                    0;

                foreach (ReferenceCoordinateImportItem importItem
                    in importItems)
                {
                    if (!importItem.IsValid)
                    {
                        invalidRowCount++;
                    }
                }

                dgReferenceCoordinateImport.Items.Refresh();

                if (invalidRowCount > 0)
                {
                    throw new InvalidOperationException(
                        $"{invalidRowCount} row(s) failed final validation.");
                }

                #endregion


                #region Insert Reference Points

                foreach (ReferenceCoordinateImportItem importItem
                    in importItems)
                {
                    int pointNameId =
                        await ResolveOrInsertPointNameAsync(
                            projectId: projectId,
                            importItem: importItem,
                            databaseConnection: databaseConnection,
                            transaction: transaction);

                    await InsertReferenceCoordinatesAsync(
                        pointNameId: pointNameId,
                        importItem: importItem,
                        databaseConnection: databaseConnection,
                        transaction: transaction);
                }

                #endregion


                #region Commit Transaction

                transaction.Commit();

                transactionCommitted =
                    true;

                return importItems.Count;

                #endregion
            }
            catch
            {
                #region Roll Back Transaction

                if (!transactionCommitted)
                {
                    try
                    {
                        transaction.Rollback();
                    }
                    catch
                    {
                        // Preserve the original exception.
                    }
                }

                throw;

                #endregion
            }
        }

        #endregion


        #region Reference Import Transaction Lock

        private static async Task AcquireReferenceImportTransactionLockAsync(
            int projectId,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Import Lock

            string lockResource =
                $"GNA_DLRreport:ReferenceImport:{projectId}";

            const string lockSql = """
        DECLARE @LockResult int;

        EXEC @LockResult = sys.sp_getapplock
            @Resource = @Resource,
            @LockMode = 'Exclusive',
            @LockOwner = 'Transaction',
            @LockTimeout = 0;

        SELECT @LockResult;
        """;

            #endregion


            #region Acquire Import Lock

            await using SqlCommand lockCommand =
                new(
                    cmdText: lockSql,
                    connection: databaseConnection,
                    transaction: transaction);

            lockCommand.Parameters.Add(
                parameterName: "@Resource",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 255)
                .Value =
                    lockResource;

            object? lockResultValue =
                await lockCommand.ExecuteScalarAsync();

            if (lockResultValue is null ||
                lockResultValue == DBNull.Value)
            {
                throw new InvalidOperationException(
                    "SQL Server did not return an import-lock result.");
            }

            int lockResult =
                Convert.ToInt32(
                    value: lockResultValue,
                    provider: CultureInfo.InvariantCulture);

            if (lockResult < 0)
            {
                throw new InvalidOperationException(
                    "Another reference-coordinate import is currently active " +
                    "for this project.");
            }

            #endregion
        }

        #endregion


        #region Resolve Or Insert Point Name

        private static async Task<int> ResolveOrInsertPointNameAsync(
            int projectId,
            ReferenceCoordinateImportItem importItem,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Validate Import Item

            if (string.IsNullOrWhiteSpace(importItem.PointName))
            {
                throw new InvalidOperationException(
                    $"Row {importItem.SourceRow}: PointName blank.");
            }

            if (string.IsNullOrWhiteSpace(importItem.ReplacementName))
            {
                throw new InvalidOperationException(
                    $"Row {importItem.SourceRow}: ReplacementName blank.");
            }

            #endregion


            #region Search Existing Point Namespace

            const string findPointSql = """
        SELECT
            PN.[PointName_ID],
            PN.[PointName],
            PN.[ReplacementName],
            PN.[IsDeleted],
            CASE
                WHEN CR.[PointName_ID] IS NULL THEN 0
                ELSE 1
            END AS [HasReference]
        FROM [dbo].[PointName] AS PN WITH (UPDLOCK, HOLDLOCK)
        LEFT JOIN [dbo].[CoordinatesReference] AS CR WITH (HOLDLOCK)
            ON CR.[PointName_ID] = PN.[PointName_ID]
                    AND CR.[IsDeleted] = 0
        WHERE
            PN.[Project_ID] = @Project_ID
            AND
            (
                UPPER(PN.[PointName]) = UPPER(@PointName)
                OR UPPER(PN.[ReplacementName]) = UPPER(@PointName)
                OR UPPER(PN.[PointName]) = UPPER(@ReplacementName)
                OR UPPER(PN.[ReplacementName]) = UPPER(@ReplacementName)
            );
        """;

            await using SqlCommand findPointCommand =
                new(
                    cmdText: findPointSql,
                    connection: databaseConnection,
                    transaction: transaction);

            findPointCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            findPointCommand.Parameters.Add(
                parameterName: "@PointName",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 50)
                .Value =
                    importItem.PointName;

            findPointCommand.Parameters.Add(
                parameterName: "@ReplacementName",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 50)
                .Value =
                    importItem.ReplacementName;

            #endregion


            #region Read Existing Point Match

            int? existingPointId =
                null;

            string existingPointName =
                string.Empty;

            string existingReplacementName =
                string.Empty;

            bool existingPointIsDeleted =
                false;

            bool existingPointHasReference =
                false;

            int matchCount =
                0;

            await using (SqlDataReader reader =
                await findPointCommand.ExecuteReaderAsync())
            {
                while (await reader.ReadAsync())
                {
                    matchCount++;

                    existingPointId =
                        reader.GetInt32(
                            i: 0);

                    existingPointName =
                        reader.GetString(
                            i: 1);

                    existingReplacementName =
                        reader.GetString(
                            i: 2);

                    existingPointIsDeleted =
                        reader.GetBoolean(
                            i: 3);

                    existingPointHasReference =
                        reader.GetInt32(
                            i: 4) == 1;
                }
            }

            #endregion


            #region Handle Multiple Namespace Matches

            if (matchCount > 1)
            {
                throw new InvalidOperationException(
                    $"'{importItem.PointName}': Name collision.");
            }

            #endregion


            #region Reuse Existing Point

            if (matchCount == 1 &&
                existingPointId.HasValue)
            {
                if (existingPointIsDeleted)
                {
                    throw new InvalidOperationException(
                        $"'{importItem.PointName}': Point is deleted.");
                }

                bool pointNameMatches =
                    string.Equals(
                        a: existingPointName,
                        b: importItem.PointName,
                        comparisonType: StringComparison.OrdinalIgnoreCase);

                bool replacementNameMatches =
                    string.Equals(
                        a: existingReplacementName,
                        b: importItem.ReplacementName,
                        comparisonType: StringComparison.OrdinalIgnoreCase);

                if (!pointNameMatches ||
                    !replacementNameMatches)
                {
                    throw new InvalidOperationException(
                        $"'{importItem.PointName}': Name mapping differs.");
                }

                if (existingPointHasReference)
                {
                    throw new InvalidOperationException(
                        $"'{importItem.PointName}': Reference coordinates already exist.");
                }

                return existingPointId.Value;
            }

            #endregion


            #region Insert New Point

            return await InsertPointNameAsync(
                projectId: projectId,
                importItem: importItem,
                databaseConnection: databaseConnection,
                transaction: transaction);

            #endregion
        }


        private static async Task<int> InsertPointNameAsync(
            int projectId,
            ReferenceCoordinateImportItem importItem,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Point Name Insert

            const string insertPointNameSql = """
        INSERT INTO [dbo].[PointName]
        (
            [PointName],
            [ReplacementName],
            [Project_ID],
            [IsDeleted]
        )
        VALUES
        (
            @PointName,
            @ReplacementName,
            @Project_ID,
            0
        );

        SELECT CAST(SCOPE_IDENTITY() AS int);
        """;

            #endregion


            #region Insert Point Name

            await using SqlCommand insertPointNameCommand =
                new(
                    cmdText: insertPointNameSql,
                    connection: databaseConnection,
                    transaction: transaction);

            insertPointNameCommand.Parameters.Add(
                parameterName: "@PointName",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 50)
                .Value =
                    importItem.PointName;

            insertPointNameCommand.Parameters.Add(
                parameterName: "@ReplacementName",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 50)
                .Value =
                    importItem.ReplacementName;

            insertPointNameCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            object? pointNameIdValue =
                await insertPointNameCommand.ExecuteScalarAsync();

            if (pointNameIdValue is null ||
                pointNameIdValue == DBNull.Value)
            {
                throw new InvalidOperationException(
                    $"'{importItem.PointName}': Point insert failed.");
            }

            return Convert.ToInt32(
                value: pointNameIdValue,
                provider: CultureInfo.InvariantCulture);

            #endregion
        }

        #endregion



        #region Insert Reference Coordinates

        private static async Task InsertReferenceCoordinatesAsync(
            int pointNameId,
            ReferenceCoordinateImportItem importItem,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Validate Coordinates

            if (!importItem.Easting.HasValue ||
                !importItem.Northing.HasValue ||
                !importItem.Height.HasValue)
            {
                throw new InvalidOperationException(
                    $"CSV row {importItem.SourceRow} does not contain " +
                    "complete valid reference coordinates.");
            }

            #endregion


            #region Define Coordinate Insert

            const string insertCoordinateSql = """
        INSERT INTO [dbo].[CoordinatesReference]
        (
            [PointName_ID],
            [Eref],
            [Nref],
            [Href]
        )
        VALUES
        (
            @PointName_ID,
            @Eref,
            @Nref,
            @Href
        );
        """;

            #endregion


            #region Insert Coordinates

            await using SqlCommand insertCoordinateCommand =
                new(
                    cmdText: insertCoordinateSql,
                    connection: databaseConnection,
                    transaction: transaction);

            insertCoordinateCommand.Parameters.Add(
                parameterName: "@PointName_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    pointNameId;


            SqlParameter eastingParameter =
                insertCoordinateCommand.Parameters.Add(
                    parameterName: "@Eref",
                    sqlDbType: System.Data.SqlDbType.Decimal);

            eastingParameter.Precision =
                18;

            eastingParameter.Scale =
                6;

            eastingParameter.Value =
                importItem.Easting.Value;


            SqlParameter northingParameter =
                insertCoordinateCommand.Parameters.Add(
                    parameterName: "@Nref",
                    sqlDbType: System.Data.SqlDbType.Decimal);

            northingParameter.Precision =
                18;

            northingParameter.Scale =
                6;

            northingParameter.Value =
                importItem.Northing.Value;


            SqlParameter heightParameter =
                insertCoordinateCommand.Parameters.Add(
                    parameterName: "@Href",
                    sqlDbType: System.Data.SqlDbType.Decimal);

            heightParameter.Precision =
                18;

            heightParameter.Scale =
                6;

            heightParameter.Value =
                importItem.Height.Value;


            int rowsInserted =
                await insertCoordinateCommand.ExecuteNonQueryAsync();

            if (rowsInserted != 1)
            {
                throw new InvalidOperationException(
                    $"Reference coordinates for '{importItem.PointName}' " +
                    "were not inserted correctly.");
            }

            #endregion
        }

        #endregion



        #endregion

        #region Geotechnical Sensor Import

        #region Geotechnical Sensor Import Models

        private sealed class GeotechSensorImportItem
        {
            public int SourceRow { get; init; }

            public string RawLine { get; init; } =
                string.Empty;

            public string SensorName { get; set; } =
                string.Empty;

            public string ReplacementName { get; set; } =
                string.Empty;

            public string SensorType { get; set; } =
                string.Empty;

            public string EastingText { get; set; } =
                string.Empty;

            public string NorthingText { get; set; } =
                string.Empty;

            public string HeightText { get; set; } =
                string.Empty;

            public double Easting { get; set; }

            public double Northing { get; set; }

            public double Height { get; set; }

            public bool IsValid { get; set; }

            public string ValidationStatus { get; set; } =
                string.Empty;
        }


        private sealed class ExistingGeotechSensorState
        {
            public int SensorID { get; init; }

            public int Project_ID { get; init; }

            public string SensorName { get; init; } =
                string.Empty;

            public string ReplacementName { get; init; } =
                string.Empty;

            public string SensorType { get; init; } =
                string.Empty;

            public bool IsDeleted { get; init; }

            public bool HasReference { get; init; }
        }

        #endregion


        #region Geotech CSV File Selection

        private async void btnSelectGeotechCsv_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Select CSV File

            OpenFileDialog openFileDialog =
                new()
                {
                    Title =
                        "Select Geotechnical Sensor CSV",

                    Filter =
                        "CSV files (*.csv)|*.csv|" +
                        "All files (*.*)|*.*",

                    DefaultExt =
                        ".csv",

                    CheckFileExists =
                        true,

                    Multiselect =
                        false
                };

            bool? fileSelected =
                openFileDialog.ShowDialog(
                    owner: this);

            if (fileSelected != true)
            {
                txtGeotechImportStatus.Text =
                    "CSV selection cancelled.";

                return;
            }

            #endregion


            #region Validate Active Project

            if (!_activeProjectId.HasValue ||
                string.IsNullOrWhiteSpace(_activeProjectName))
            {
                txtGeotechImportStatus.Text =
                    "Select an active project.";

                MessageBox.Show(
                    owner: this,
                    messageBoxText:
                        "Select an active project before importing geotechnical sensors.",
                    caption:
                        "Active Project Required",
                    button:
                        MessageBoxButton.OK,
                    icon:
                        MessageBoxImage.Warning);

                return;
            }

            #endregion


            #region Capture Import Project Context

            _geotechImportProjectId =
                _activeProjectId.Value;

            _geotechImportProjectName =
                _activeProjectName;

            #endregion


            #region Store Selected CSV

            _selectedGeotechCsvPath =
                openFileDialog.FileName;

            txtGeotechCsvPath.Text =
                _selectedGeotechCsvPath;

            #endregion


            #region Parse And Validate CSV

            await ParseAndValidateSelectedGeotechCsvAsync();

            #endregion
        }


        private void btnClearGeotechImport_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Clear Geotech Import State

            ResetGeotechImportState(
                statusMessage:
                    "No CSV selected.");

            #endregion
        }

        #endregion


        #region Geotech Import State Reset

        private void ResetGeotechImportState(
            string statusMessage)
        {
            #region Clear Captured Import Context

            _geotechImportProjectId =
                null;

            _geotechImportProjectName =
                string.Empty;

            _selectedGeotechCsvPath =
                string.Empty;

            #endregion


            #region Clear User Interface

            txtGeotechCsvPath.Text =
                string.Empty;

            _geotechImportItems.Clear();

            dgGeotechSensorImport.SelectedItem =
                null;

            btnCommitGeotechImport.IsEnabled =
                false;

            txtGeotechImportStatus.Text =
                statusMessage;

            #endregion
        }

        #endregion


        #region Geotech Import Preview Selection

        private void dgGeotechSensorImport_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Display Selected Row Validation

            if (dgGeotechSensorImport.SelectedItem
                is not GeotechSensorImportItem selectedItem)
            {
                return;
            }

            txtGeotechImportStatus.Text =
                selectedItem.ValidationStatus;

            #endregion
        }

        #endregion


        #region Parse And Validate Selected Geotech CSV

        private async Task ParseAndValidateSelectedGeotechCsvAsync()
        {
            #region Validate Import Context

            if (!_geotechImportProjectId.HasValue)
            {
                throw new InvalidOperationException(
                    "The Geotech import does not have a target Project_ID.");
            }

            if (string.IsNullOrWhiteSpace(_geotechImportProjectName))
            {
                throw new InvalidOperationException(
                    "The Geotech import does not have a target project name.");
            }

            if (string.IsNullOrWhiteSpace(_selectedGeotechCsvPath))
            {
                throw new InvalidOperationException(
                    "No Geotech CSV has been selected.");
            }

            if (!File.Exists(
                path: _selectedGeotechCsvPath))
            {
                throw new FileNotFoundException(
                    message:
                        "The selected Geotech CSV no longer exists.",
                    fileName:
                        _selectedGeotechCsvPath);
            }

            #endregion


            #region Prepare User Interface

            btnSelectGeotechCsv.IsEnabled =
                false;

            btnClearGeotechImport.IsEnabled =
                false;

            btnCommitGeotechImport.IsEnabled =
                false;

            txtGeotechImportStatus.Text =
                $"Validating CSV for project '{_geotechImportProjectName}'...";

            #endregion


            try
            {
                #region Parse Source CSV

                List<GeotechSensorImportItem> importItems =
                    ParseGeotechSensorCsv(
                        csvPath:
                            _selectedGeotechCsvPath);

                #endregion


                #region Load Existing Geotechnical Sensors

                List<ExistingGeotechSensorState> existingSensors =
                    await LoadExistingGeotechSensorStatesAsync(
                        projectId:
                            _geotechImportProjectId.Value,
                        databaseConnection:
                            null,
                        transaction:
                            null,
                        applyCommitLocks:
                            false);

                #endregion


                #region Validate Import

                ValidateGeotechSensorImport(
                    importItems:
                        importItems,
                    existingSensors:
                        existingSensors);

                #endregion


                #region Populate Preview DataGrid

                _geotechImportItems.Clear();

                foreach (GeotechSensorImportItem importItem
                    in importItems)
                {
                    _geotechImportItems.Add(
                        item:
                            importItem);
                }

                #endregion


                #region Determine Validation Result

                int invalidRowCount =
                    importItems.Count(
                        predicate:
                            item => !item.IsValid);

                if (importItems.Count == 0)
                {
                    txtGeotechImportStatus.Text =
                        $"Project '{_geotechImportProjectName}': " +
                        "the selected CSV contains no records.";

                    return;
                }

                if (invalidRowCount > 0)
                {
                    txtGeotechImportStatus.Text =
                        $"Project '{_geotechImportProjectName}': " +
                        $"{invalidRowCount} of {importItems.Count} row(s) contain errors. " +
                        "Import blocked.";

                    return;
                }

                btnCommitGeotechImport.IsEnabled =
                    true;

                txtGeotechImportStatus.Text =
                    $"Project '{_geotechImportProjectName}': " +
                    $"{importItems.Count} sensor(s) valid and ready to import.";

                #endregion
            }
            catch (Exception ex)
            {
                #region Report Validation Failure

                _geotechImportItems.Clear();

                btnCommitGeotechImport.IsEnabled =
                    false;

                txtGeotechImportStatus.Text =
                    $"CSV validation failed: {ex.Message}";

                #endregion
            }
            finally
            {
                #region Restore Controls

                btnSelectGeotechCsv.IsEnabled =
                    true;

                btnClearGeotechImport.IsEnabled =
                    true;

                #endregion
            }
        }

        #endregion


        #region Geotech CSV Parsing

        private static List<GeotechSensorImportItem>
            ParseGeotechSensorCsv(
                string csvPath)
        {
            #region Validate CSV Path

            if (string.IsNullOrWhiteSpace(csvPath))
            {
                throw new ArgumentException(
                    message:
                        "CSV path cannot be empty.",
                    paramName:
                        nameof(csvPath));
            }

            #endregion


            #region Read CSV Lines

            string[] sourceLines =
                File.ReadAllLines(
                    path:
                        csvPath);

            List<GeotechSensorImportItem> importItems =
                new();

            #endregion


            #region Parse CSV Rows

            for (int lineIndex = 0;
                 lineIndex < sourceLines.Length;
                 lineIndex++)
            {
                int sourceRow =
                    lineIndex + 1;

                string rawLine =
                    sourceLines[lineIndex];

                GeotechSensorImportItem importItem =
                    new()
                    {
                        SourceRow = sourceRow,
                        RawLine = rawLine,
                        IsValid = true,
                        ValidationStatus = "Valid"
                    };


                #region Parse CSV Fields

                if (!TryParseCsvLine(
                    line:
                        rawLine,
                    fields:
                        out List<string> fields,
                    errorMessage:
                        out string csvError))
                {
                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            csvError);

                    importItems.Add(
                        item:
                            importItem);

                    continue;
                }

                #endregion


                #region Validate Column Count

                if (fields.Count != 6)
                {
                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            $"Expected 6 columns; found {fields.Count}.");

                    importItems.Add(
                        item:
                            importItem);

                    continue;
                }

                #endregion


                #region Extract Sensor Fields

                importItem.SensorName =
                    fields[0].Trim();

                importItem.EastingText =
                    fields[1].Trim();

                importItem.NorthingText =
                    fields[2].Trim();

                importItem.HeightText =
                    fields[3].Trim();

                string replacementName =
                    fields[4].Trim();

                importItem.ReplacementName =
                    string.IsNullOrWhiteSpace(replacementName)
                        ? importItem.SensorName
                        : replacementName;

                importItem.SensorType =
                    fields[5].Trim();

                #endregion


                #region Validate Sensor Name

                if (string.IsNullOrWhiteSpace(importItem.SensorName))
                {
                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            "Sensor Name is blank.");
                }
                else if (importItem.SensorName.Length > 50)
                {
                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            "Sensor Name exceeds 50 characters.");
                }

                #endregion


                #region Validate Replacement Name

                if (string.IsNullOrWhiteSpace(importItem.ReplacementName))
                {
                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            "Replacement Name is blank.");
                }
                else if (importItem.ReplacementName.Length > 50)
                {
                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            "Replacement Name exceeds 50 characters.");
                }

                #endregion


                #region Validate Sensor Type

                if (string.IsNullOrWhiteSpace(importItem.SensorType))
                {
                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            "Sensor Type is blank.");
                }
                else if (importItem.SensorType.Length > 50)
                {
                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            "Sensor Type exceeds 50 characters.");
                }

                #endregion


                #region Parse Optional Reference Coordinates

                bool eastingBlank =
                    string.IsNullOrWhiteSpace(importItem.EastingText);

                bool northingBlank =
                    string.IsNullOrWhiteSpace(importItem.NorthingText);

                bool heightBlank =
                    string.IsNullOrWhiteSpace(importItem.HeightText);

                if (eastingBlank &&
                    northingBlank &&
                    heightBlank)
                {
                    importItem.Easting =
                        0.0;

                    importItem.Northing =
                        0.0;

                    importItem.Height =
                        0.0;

                    importItem.EastingText =
                        "0.000";

                    importItem.NorthingText =
                        "0.000";

                    importItem.HeightText =
                        "0.000";
                }
                else if (eastingBlank ||
                         northingBlank ||
                         heightBlank)
                {
                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            "E, N and Ht must all be supplied or all blank.");
                }
                else
                {
                    ParseGeotechCoordinate(
                        coordinateText:
                            importItem.EastingText,
                        coordinateName:
                            "Easting",
                        importItem:
                            importItem,
                        valueSetter:
                            value => importItem.Easting = value);

                    ParseGeotechCoordinate(
                        coordinateText:
                            importItem.NorthingText,
                        coordinateName:
                            "Northing",
                        importItem:
                            importItem,
                        valueSetter:
                            value => importItem.Northing = value);

                    ParseGeotechCoordinate(
                        coordinateText:
                            importItem.HeightText,
                        coordinateName:
                            "Height",
                        importItem:
                            importItem,
                        valueSetter:
                            value => importItem.Height = value);

                    if (importItem.IsValid)
                    {
                        importItem.Easting =
                            Math.Round(
                                value:
                                    importItem.Easting,
                                digits:
                                    6,
                                mode:
                                    MidpointRounding.AwayFromZero);

                        importItem.Northing =
                            Math.Round(
                                value:
                                    importItem.Northing,
                                digits:
                                    6,
                                mode:
                                    MidpointRounding.AwayFromZero);

                        importItem.Height =
                            Math.Round(
                                value:
                                    importItem.Height,
                                digits:
                                    6,
                                mode:
                                    MidpointRounding.AwayFromZero);

                        importItem.EastingText =
                            importItem.Easting.ToString(
                                format:
                                    "F3",
                                provider:
                                    CultureInfo.InvariantCulture);

                        importItem.NorthingText =
                            importItem.Northing.ToString(
                                format:
                                    "F3",
                                provider:
                                    CultureInfo.InvariantCulture);

                        importItem.HeightText =
                            importItem.Height.ToString(
                                format:
                                    "F3",
                                provider:
                                    CultureInfo.InvariantCulture);
                    }
                }

                #endregion


                #region Add Parsed Row

                importItems.Add(
                    item:
                        importItem);

                #endregion
            }

            #endregion


            #region Return Parsed CSV

            return importItems;

            #endregion
        }


        private static void ParseGeotechCoordinate(
            string coordinateText,
            string coordinateName,
            GeotechSensorImportItem importItem,
            Action<double> valueSetter)
        {
            #region Parse Coordinate

            if (!double.TryParse(
                s:
                    coordinateText,
                style:
                    NumberStyles.Float,
                provider:
                    CultureInfo.InvariantCulture,
                result:
                    out double coordinateValue) ||
                !double.IsFinite(coordinateValue))
            {
                AppendGeotechValidationError(
                    importItem:
                        importItem,
                    errorMessage:
                        $"Invalid {coordinateName} '{coordinateText}'.");

                return;
            }

            #endregion


            #region Validate Decimal 18,6 Range

            const double maximumCoordinateMagnitude =
                999999999999.999999;

            if (Math.Abs(coordinateValue) >
                maximumCoordinateMagnitude)
            {
                AppendGeotechValidationError(
                    importItem:
                        importItem,
                    errorMessage:
                        $"{coordinateName} exceeds decimal(18,6) range.");

                return;
            }

            #endregion


            #region Store Coordinate

            valueSetter(
                obj:
                    coordinateValue);

            #endregion
        }

        #endregion


        #region Load Existing Geotechnical Sensors

        private async Task<List<ExistingGeotechSensorState>>
            LoadExistingGeotechSensorStatesAsync(
                int projectId,
                SqlConnection? databaseConnection,
                SqlTransaction? transaction,
                bool applyCommitLocks)
        {
            #region Validate Project ID

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName:
                        nameof(projectId),
                    message:
                        "Project_ID must be greater than zero.");
            }

            #endregion


            #region Open Connection When Required

            bool ownsConnection =
                databaseConnection is null;

            SqlConnection connection =
                databaseConnection
                ?? new SqlConnection(
                    connectionString:
                        GetTrackGeometryConnectionString());

            if (ownsConnection)
            {
                await connection.OpenAsync();
            }

            #endregion


            try
            {
                #region Define Existing Sensor Query

                string sensorLockHint =
                    applyCommitLocks
                        ? " WITH (UPDLOCK, HOLDLOCK)"
                        : string.Empty;

                string referenceLockHint =
                    applyCommitLocks
                        ? " WITH (UPDLOCK, HOLDLOCK)"
                        : string.Empty;

                string existingSensorSql =
                    $"""
                    SELECT
                        GS.[SensorID],
                        GS.[Project_ID],
                        GS.[SensorName],
                        GS.[ReplacementName],
                        GS.[SensorType],
                        GS.[IsDeleted],
                        CASE
                            WHEN CR.[Coordinate_ID] IS NULL THEN 0
                            ELSE 1
                        END AS [HasReference]
                    FROM [dbo].[GeotecSensors] AS GS{sensorLockHint}
                    LEFT JOIN [dbo].[CoordinatesReference] AS CR{referenceLockHint}
                        ON CR.[SensorID] = GS.[SensorID]
                        AND CR.[IsDeleted] = 0
                    WHERE
                        GS.[Project_ID] = @Project_ID;
                    """;

                #endregion


                #region Execute Existing Sensor Query

                List<ExistingGeotechSensorState> existingSensors =
                    new();

                await using SqlCommand sensorCommand =
                    transaction is null
                        ? new SqlCommand(
                            cmdText:
                                existingSensorSql,
                            connection:
                                connection)
                        : new SqlCommand(
                            cmdText:
                                existingSensorSql,
                            connection:
                                connection,
                            transaction:
                                transaction);

                sensorCommand.Parameters.Add(
                    parameterName:
                        "@Project_ID",
                    sqlDbType:
                        System.Data.SqlDbType.Int)
                    .Value =
                        projectId;

                await using SqlDataReader reader =
                    await sensorCommand.ExecuteReaderAsync();

                while (await reader.ReadAsync())
                {
                    existingSensors.Add(
                        item:
                            new ExistingGeotechSensorState
                            {
                                SensorID =
                                    reader.GetInt32(0),

                                Project_ID =
                                    reader.GetInt32(1),

                                SensorName =
                                    reader.GetString(2),

                                ReplacementName =
                                    reader.GetString(3),

                                SensorType =
                                    reader.GetString(4),

                                IsDeleted =
                                    reader.GetBoolean(5),

                                HasReference =
                                    reader.GetInt32(6) == 1
                            });
                }

                return existingSensors;

                #endregion
            }
            finally
            {
                #region Dispose Owned Connection

                if (ownsConnection)
                {
                    await connection.DisposeAsync();
                }

                #endregion
            }
        }

        #endregion


        #region Geotech Import Validation

        private static void ValidateGeotechSensorImport(
            List<GeotechSensorImportItem> importItems,
            List<ExistingGeotechSensorState> existingSensors)
        {
            #region Validate Arguments

            ArgumentNullException.ThrowIfNull(
                argument:
                    importItems);

            ArgumentNullException.ThrowIfNull(
                argument:
                    existingSensors);

            #endregion


            #region Validate Combined Names Within CSV

            Dictionary<string, GeotechSensorImportItem> incomingNames =
                new(
                    comparer:
                        StringComparer.OrdinalIgnoreCase);

            foreach (GeotechSensorImportItem importItem
                in importItems)
            {
                RegisterGeotechImportName(
                    name:
                        importItem.SensorName,
                    importItem:
                        importItem,
                    incomingNames:
                        incomingNames);

                if (!string.Equals(
                    a:
                        importItem.SensorName,
                    b:
                        importItem.ReplacementName,
                    comparisonType:
                        StringComparison.OrdinalIgnoreCase))
                {
                    RegisterGeotechImportName(
                        name:
                            importItem.ReplacementName,
                        importItem:
                            importItem,
                        incomingNames:
                            incomingNames);
                }
            }

            #endregion


            #region Build Existing Sensor Namespace

            Dictionary<string, List<ExistingGeotechSensorState>> existingNames =
                new(
                    comparer:
                        StringComparer.OrdinalIgnoreCase);

            foreach (ExistingGeotechSensorState existingSensor
                in existingSensors)
            {
                AddExistingGeotechName(
                    name:
                        existingSensor.SensorName,
                    sensor:
                        existingSensor,
                    existingNames:
                        existingNames);

                if (!string.Equals(
                    a:
                        existingSensor.SensorName,
                    b:
                        existingSensor.ReplacementName,
                    comparisonType:
                        StringComparison.OrdinalIgnoreCase))
                {
                    AddExistingGeotechName(
                        name:
                            existingSensor.ReplacementName,
                        sensor:
                            existingSensor,
                        existingNames:
                            existingNames);
                }
            }

            #endregion


            #region Validate Against Existing Database Sensors

            foreach (GeotechSensorImportItem importItem
                in importItems)
            {
                HashSet<int> matchedSensorIds =
                    new();

                AddGeotechExistingMatches(
                    name:
                        importItem.SensorName,
                    existingNames:
                        existingNames,
                    matchedSensorIds:
                        matchedSensorIds);

                AddGeotechExistingMatches(
                    name:
                        importItem.ReplacementName,
                    existingNames:
                        existingNames,
                    matchedSensorIds:
                        matchedSensorIds);

                if (matchedSensorIds.Count == 0)
                {
                    continue;
                }

                if (matchedSensorIds.Count > 1)
                {
                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            "Sensor/Replacement Name conflicts with multiple database sensors.");

                    continue;
                }

                int matchedSensorId =
                    matchedSensorIds.Single();

                ExistingGeotechSensorState existingSensor =
                    existingSensors.Single(
                        predicate:
                            sensor => sensor.SensorID == matchedSensorId);

                bool exactMapping =
                    string.Equals(
                        a:
                            importItem.SensorName,
                        b:
                            existingSensor.SensorName,
                        comparisonType:
                            StringComparison.OrdinalIgnoreCase)
                    &&
                    string.Equals(
                        a:
                            importItem.ReplacementName,
                        b:
                            existingSensor.ReplacementName,
                        comparisonType:
                            StringComparison.OrdinalIgnoreCase)
                    &&
                    string.Equals(
                        a:
                            importItem.SensorType,
                        b:
                            existingSensor.SensorType,
                        comparisonType:
                            StringComparison.OrdinalIgnoreCase);

                if (exactMapping &&
                    !existingSensor.IsDeleted &&
                    !existingSensor.HasReference)
                {
                    continue;
                }

                AppendGeotechValidationError(
                    importItem:
                        importItem,
                    errorMessage:
                        "Sensor or Replacement Name already exists in this project.");
            }

            #endregion
        }


        private static void RegisterGeotechImportName(
            string name,
            GeotechSensorImportItem importItem,
            Dictionary<string, GeotechSensorImportItem> incomingNames)
        {
            #region Ignore Blank Names Already Reported

            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            #endregion


            #region Register Or Report Duplicate

            if (incomingNames.TryGetValue(
                key:
                    name,
                value:
                    out GeotechSensorImportItem? existingItem))
            {
                if (!ReferenceEquals(
                    objA:
                        existingItem,
                    objB:
                        importItem))
                {
                    AppendGeotechValidationError(
                        importItem:
                            existingItem,
                        errorMessage:
                            $"Duplicate name '{name}' in CSV.");

                    AppendGeotechValidationError(
                        importItem:
                            importItem,
                        errorMessage:
                            $"Duplicate name '{name}' in CSV.");
                }

                return;
            }

            incomingNames.Add(
                key:
                    name,
                value:
                    importItem);

            #endregion
        }


        private static void AddExistingGeotechName(
            string name,
            ExistingGeotechSensorState sensor,
            Dictionary<string, List<ExistingGeotechSensorState>> existingNames)
        {
            #region Register Existing Name

            if (!existingNames.TryGetValue(
                key:
                    name,
                value:
                    out List<ExistingGeotechSensorState>? sensors))
            {
                sensors =
                    new List<ExistingGeotechSensorState>();

                existingNames.Add(
                    key:
                        name,
                    value:
                        sensors);
            }

            sensors.Add(
                item:
                    sensor);

            #endregion
        }


        private static void AddGeotechExistingMatches(
            string name,
            Dictionary<string, List<ExistingGeotechSensorState>> existingNames,
            HashSet<int> matchedSensorIds)
        {
            #region Add Matching Sensor IDs

            if (string.IsNullOrWhiteSpace(name) ||
                !existingNames.TryGetValue(
                    key:
                        name,
                    value:
                        out List<ExistingGeotechSensorState>? sensors))
            {
                return;
            }

            foreach (ExistingGeotechSensorState sensor
                in sensors)
            {
                matchedSensorIds.Add(
                    item:
                        sensor.SensorID);
            }

            #endregion
        }


        private static void AppendGeotechValidationError(
            GeotechSensorImportItem importItem,
            string errorMessage)
        {
            #region Append Validation Error

            importItem.IsValid =
                false;

            if (string.IsNullOrWhiteSpace(importItem.ValidationStatus) ||
                string.Equals(
                    a:
                        importItem.ValidationStatus,
                    b:
                        "Valid",
                    comparisonType:
                        StringComparison.Ordinal))
            {
                importItem.ValidationStatus =
                    errorMessage;

                return;
            }

            if (!importItem.ValidationStatus.Contains(
                value:
                    errorMessage,
                comparisonType:
                    StringComparison.Ordinal))
            {
                importItem.ValidationStatus +=
                    $" {errorMessage}";
            }

            #endregion
        }

        #endregion


        #region Commit Geotechnical Sensor Import

        private async void btnCommitGeotechImport_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Commit Context

            if (!_geotechImportProjectId.HasValue)
            {
                btnCommitGeotechImport.IsEnabled =
                    false;

                txtGeotechImportStatus.Text =
                    "Import blocked: no target project.";

                return;
            }

            if (!_activeProjectId.HasValue ||
                _activeProjectId.Value != _geotechImportProjectId.Value)
            {
                btnCommitGeotechImport.IsEnabled =
                    false;

                txtGeotechImportStatus.Text =
                    "Import blocked: active project changed. Select the CSV again.";

                return;
            }

            if (_geotechImportItems.Count == 0 ||
                _geotechImportItems.Any(
                    predicate:
                        item => !item.IsValid))
            {
                btnCommitGeotechImport.IsEnabled =
                    false;

                txtGeotechImportStatus.Text =
                    "Import blocked: preview contains validation errors or no sensors.";

                return;
            }

            #endregion


            #region Confirm Import

            MessageBoxResult confirmationResult =
                MessageBox.Show(
                    owner:
                        this,
                    messageBoxText:
                        $"Project: {_geotechImportProjectName}\n\n" +
                        $"Sensors to import: {_geotechImportItems.Count}\n" +
                        "Validation errors: 0\n\n" +
                        "Commit this import?",
                    caption:
                        "Confirm Geotechnical Sensor Import",
                    button:
                        MessageBoxButton.YesNo,
                    icon:
                        MessageBoxImage.Question);

            if (confirmationResult != MessageBoxResult.Yes)
            {
                txtGeotechImportStatus.Text =
                    "Import cancelled. No data was written to the database.";

                return;
            }

            #endregion


            #region Prepare Commit

            btnCommitGeotechImport.IsEnabled =
                false;

            btnSelectGeotechCsv.IsEnabled =
                false;

            btnClearGeotechImport.IsEnabled =
                false;

            txtGeotechImportStatus.Text =
                $"Revalidating and importing {_geotechImportItems.Count} sensor(s)...";

            #endregion


            try
            {
                #region Execute Atomic Import

                int importedSensorCount =
                    await CommitGeotechSensorImportAsync(
                        projectId:
                            _geotechImportProjectId.Value,
                        projectName:
                            _geotechImportProjectName);

                #endregion


                #region Report Successful Import

                btnCommitGeotechImport.IsEnabled =
                    false;

                txtGeotechImportStatus.Text =
                    $"{importedSensorCount} sensor(s) imported successfully " +
                    $"into project '{_geotechImportProjectName}'.";

                #endregion
            }
            catch (Exception ex)
            {
                #region Report Failed Import

                bool previewStillValid =
                    !_geotechImportItems.Any(
                        predicate:
                            item => !item.IsValid);

                btnCommitGeotechImport.IsEnabled =
                    previewStillValid;

                dgGeotechSensorImport.Items.Refresh();

                txtGeotechImportStatus.Text =
                    $"Import failed. No data committed. {ex.Message}";

                #endregion
            }
            finally
            {
                #region Restore Import Controls

                btnSelectGeotechCsv.IsEnabled =
                    true;

                btnClearGeotechImport.IsEnabled =
                    true;

                #endregion
            }
        }

        #endregion


        #region Atomic Geotechnical Sensor Database Import

        private async Task<int> CommitGeotechSensorImportAsync(
            int projectId,
            string projectName)
        {
            #region Validate Parameters

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName:
                        nameof(projectId),
                    message:
                        "Project_ID must be greater than zero.");
            }

            if (string.IsNullOrWhiteSpace(projectName))
            {
                throw new ArgumentException(
                    message:
                        "Project name cannot be blank.",
                    paramName:
                        nameof(projectName));
            }

            #endregion


            #region Capture Preview Dataset

            List<GeotechSensorImportItem> importItems =
                new(
                    collection:
                        _geotechImportItems);

            if (importItems.Count == 0)
            {
                throw new InvalidOperationException(
                    "The Geotech preview contains no sensors.");
            }

            #endregion


            #region Open Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        GetTrackGeometryConnectionString());

            await databaseConnection.OpenAsync();

            #endregion


            #region Begin Serializable Transaction

            using SqlTransaction transaction =
                databaseConnection.BeginTransaction(
                    iso:
                        System.Data.IsolationLevel.Serializable);

            bool transactionCommitted =
                false;

            #endregion


            try
            {
                #region Acquire Geotech Import Lock

                await AcquireGeotechImportTransactionLockAsync(
                    projectId:
                        projectId,
                    databaseConnection:
                        databaseConnection,
                    transaction:
                        transaction);

                #endregion


                #region Revalidate Target Project

                await ValidateImportTargetProjectAsync(
                    projectId:
                        projectId,
                    expectedProjectName:
                        projectName,
                    databaseConnection:
                        databaseConnection,
                    transaction:
                        transaction);

                #endregion


                #region Reset Preview Validation Status

                foreach (GeotechSensorImportItem importItem
                    in importItems)
                {
                    importItem.IsValid =
                        true;

                    importItem.ValidationStatus =
                        "Valid";
                }

                #endregion


                #region Reload Current Database Sensors

                List<ExistingGeotechSensorState> existingSensors =
                    await LoadExistingGeotechSensorStatesAsync(
                        projectId:
                            projectId,
                        databaseConnection:
                            databaseConnection,
                        transaction:
                            transaction,
                        applyCommitLocks:
                            true);

                #endregion


                #region Revalidate Import Against Database

                ValidateGeotechSensorImport(
                    importItems:
                        importItems,
                    existingSensors:
                        existingSensors);

                dgGeotechSensorImport.Items.Refresh();

                int invalidRowCount =
                    importItems.Count(
                        predicate:
                            item => !item.IsValid);

                if (invalidRowCount > 0)
                {
                    throw new InvalidOperationException(
                        $"{invalidRowCount} row(s) failed final validation.");
                }

                #endregion


                #region Insert Sensors And Reference Coordinates

                foreach (GeotechSensorImportItem importItem
                    in importItems)
                {
                    int sensorId =
                        await ResolveOrInsertGeotechSensorAsync(
                            projectId:
                                projectId,
                            importItem:
                                importItem,
                            databaseConnection:
                                databaseConnection,
                            transaction:
                                transaction);

                    await InsertGeotechReferenceCoordinatesAsync(
                        sensorId:
                            sensorId,
                        importItem:
                            importItem,
                        databaseConnection:
                            databaseConnection,
                        transaction:
                            transaction);
                }

                #endregion


                #region Commit Transaction

                transaction.Commit();

                transactionCommitted =
                    true;

                return importItems.Count;

                #endregion
            }
            catch
            {
                #region Roll Back Transaction

                if (!transactionCommitted)
                {
                    try
                    {
                        transaction.Rollback();
                    }
                    catch
                    {
                        // Preserve the original exception.
                    }
                }

                throw;

                #endregion
            }
        }

        #endregion


        #region Geotech Import Transaction Lock

        private static async Task AcquireGeotechImportTransactionLockAsync(
            int projectId,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Geotech Import Lock

            string lockResource =
                $"GNA_DLRreport:GeotechImport:{projectId}";

            const string lockSql = """
                DECLARE @LockResult int;

                EXEC @LockResult = sys.sp_getapplock
                    @Resource = @Resource,
                    @LockMode = 'Exclusive',
                    @LockOwner = 'Transaction',
                    @LockTimeout = 0;

                SELECT @LockResult;
                """;

            #endregion


            #region Acquire Geotech Import Lock

            await using SqlCommand lockCommand =
                new(
                    cmdText:
                        lockSql,
                    connection:
                        databaseConnection,
                    transaction:
                        transaction);

            lockCommand.Parameters.Add(
                parameterName:
                    "@Resource",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    255)
                .Value =
                    lockResource;

            object? lockResultValue =
                await lockCommand.ExecuteScalarAsync();

            if (lockResultValue is null ||
                lockResultValue == DBNull.Value)
            {
                throw new InvalidOperationException(
                    "SQL Server returned no Geotech import-lock result.");
            }

            int lockResult =
                Convert.ToInt32(
                    value:
                        lockResultValue,
                    provider:
                        CultureInfo.InvariantCulture);

            if (lockResult < 0)
            {
                throw new InvalidOperationException(
                    "Another Geotech import is currently active for this project.");
            }

            #endregion
        }

        #endregion


        #region Resolve Or Insert Geotechnical Sensor

        private static async Task<int> ResolveOrInsertGeotechSensorAsync(
            int projectId,
            GeotechSensorImportItem importItem,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Search Existing Sensor Namespace

            const string findSensorSql = """
                SELECT
                    GS.[SensorID],
                    GS.[SensorName],
                    GS.[ReplacementName],
                    GS.[SensorType],
                    GS.[IsDeleted],
                    CASE
                        WHEN CR.[Coordinate_ID] IS NULL THEN 0
                        ELSE 1
                    END AS [HasReference]
                FROM [dbo].[GeotecSensors] AS GS WITH (UPDLOCK, HOLDLOCK)
                LEFT JOIN [dbo].[CoordinatesReference] AS CR WITH (UPDLOCK, HOLDLOCK)
                    ON CR.[SensorID] = GS.[SensorID]
                    AND CR.[IsDeleted] = 0
                WHERE
                    GS.[Project_ID] = @Project_ID
                    AND
                    (
                        UPPER(GS.[SensorName]) = UPPER(@SensorName)
                        OR UPPER(GS.[ReplacementName]) = UPPER(@SensorName)
                        OR UPPER(GS.[SensorName]) = UPPER(@ReplacementName)
                        OR UPPER(GS.[ReplacementName]) = UPPER(@ReplacementName)
                    );
                """;

            await using SqlCommand findSensorCommand =
                new(
                    cmdText:
                        findSensorSql,
                    connection:
                        databaseConnection,
                    transaction:
                        transaction);

            findSensorCommand.Parameters.Add(
                parameterName:
                    "@Project_ID",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            findSensorCommand.Parameters.Add(
                parameterName:
                    "@SensorName",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    50)
                .Value =
                    importItem.SensorName;

            findSensorCommand.Parameters.Add(
                parameterName:
                    "@ReplacementName",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    50)
                .Value =
                    importItem.ReplacementName;

            int? matchingSensorId =
                null;

            string matchingSensorName =
                string.Empty;

            string matchingReplacementName =
                string.Empty;

            string matchingSensorType =
                string.Empty;

            bool matchingIsDeleted =
                false;

            bool matchingHasReference =
                false;

            int matchCount =
                0;

            await using (SqlDataReader reader =
                await findSensorCommand.ExecuteReaderAsync())
            {
                while (await reader.ReadAsync())
                {
                    matchCount++;

                    if (matchCount == 1)
                    {
                        matchingSensorId =
                            reader.GetInt32(0);

                        matchingSensorName =
                            reader.GetString(1);

                        matchingReplacementName =
                            reader.GetString(2);

                        matchingSensorType =
                            reader.GetString(3);

                        matchingIsDeleted =
                            reader.GetBoolean(4);

                        matchingHasReference =
                            reader.GetInt32(5) == 1;
                    }
                }
            }

            #endregion


            #region Reuse Exact Sensor Without Reference Coordinates

            if (matchCount == 1 &&
                matchingSensorId.HasValue &&
                !matchingIsDeleted &&
                !matchingHasReference &&
                string.Equals(
                    a:
                        matchingSensorName,
                    b:
                        importItem.SensorName,
                    comparisonType:
                        StringComparison.OrdinalIgnoreCase) &&
                string.Equals(
                    a:
                        matchingReplacementName,
                    b:
                        importItem.ReplacementName,
                    comparisonType:
                        StringComparison.OrdinalIgnoreCase) &&
                string.Equals(
                    a:
                        matchingSensorType,
                    b:
                        importItem.SensorType,
                    comparisonType:
                        StringComparison.OrdinalIgnoreCase))
            {
                return matchingSensorId.Value;
            }

            if (matchCount > 0)
            {
                throw new InvalidOperationException(
                    $"Sensor '{importItem.SensorName}' conflicts with an existing sensor.");
            }

            #endregion


            #region Insert New Geotechnical Sensor

            const string insertSensorSql = """
                INSERT INTO [dbo].[GeotecSensors]
                (
                    [SensorName],
                    [ReplacementName],
                    [SensorType],
                    [Project_ID]
                )
                OUTPUT INSERTED.[SensorID]
                VALUES
                (
                    @SensorName,
                    @ReplacementName,
                    @SensorType,
                    @Project_ID
                );
                """;

            await using SqlCommand insertSensorCommand =
                new(
                    cmdText:
                        insertSensorSql,
                    connection:
                        databaseConnection,
                    transaction:
                        transaction);

            insertSensorCommand.Parameters.Add(
                parameterName:
                    "@SensorName",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    50)
                .Value =
                    importItem.SensorName;

            insertSensorCommand.Parameters.Add(
                parameterName:
                    "@ReplacementName",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    50)
                .Value =
                    importItem.ReplacementName;

            insertSensorCommand.Parameters.Add(
                parameterName:
                    "@SensorType",
                sqlDbType:
                    System.Data.SqlDbType.NVarChar,
                size:
                    50)
                .Value =
                    importItem.SensorType;

            insertSensorCommand.Parameters.Add(
                parameterName:
                    "@Project_ID",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            object? sensorIdValue =
                await insertSensorCommand.ExecuteScalarAsync();

            if (sensorIdValue is null ||
                sensorIdValue == DBNull.Value)
            {
                throw new InvalidOperationException(
                    $"Sensor '{importItem.SensorName}' insert failed.");
            }

            return Convert.ToInt32(
                value:
                    sensorIdValue,
                provider:
                    CultureInfo.InvariantCulture);

            #endregion
        }

        #endregion


        #region Insert Geotechnical Reference Coordinates

        private static async Task InsertGeotechReferenceCoordinatesAsync(
            int sensorId,
            GeotechSensorImportItem importItem,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Coordinate Insert

            const string insertCoordinateSql = """
                INSERT INTO [dbo].[CoordinatesReference]
                (
                    [SensorID],
                    [Eref],
                    [Nref],
                    [Href]
                )
                VALUES
                (
                    @SensorID,
                    @Eref,
                    @Nref,
                    @Href
                );
                """;

            #endregion


            #region Insert Coordinates

            await using SqlCommand insertCoordinateCommand =
                new(
                    cmdText:
                        insertCoordinateSql,
                    connection:
                        databaseConnection,
                    transaction:
                        transaction);

            insertCoordinateCommand.Parameters.Add(
                parameterName:
                    "@SensorID",
                sqlDbType:
                    System.Data.SqlDbType.Int)
                .Value =
                    sensorId;

            SqlParameter eastingParameter =
                insertCoordinateCommand.Parameters.Add(
                    parameterName:
                        "@Eref",
                    sqlDbType:
                        System.Data.SqlDbType.Decimal);

            eastingParameter.Precision =
                18;

            eastingParameter.Scale =
                6;

            eastingParameter.Value =
                Convert.ToDecimal(
                    value:
                        importItem.Easting,
                    provider:
                        CultureInfo.InvariantCulture);

            SqlParameter northingParameter =
                insertCoordinateCommand.Parameters.Add(
                    parameterName:
                        "@Nref",
                    sqlDbType:
                        System.Data.SqlDbType.Decimal);

            northingParameter.Precision =
                18;

            northingParameter.Scale =
                6;

            northingParameter.Value =
                Convert.ToDecimal(
                    value:
                        importItem.Northing,
                    provider:
                        CultureInfo.InvariantCulture);

            SqlParameter heightParameter =
                insertCoordinateCommand.Parameters.Add(
                    parameterName:
                        "@Href",
                    sqlDbType:
                        System.Data.SqlDbType.Decimal);

            heightParameter.Precision =
                18;

            heightParameter.Scale =
                6;

            heightParameter.Value =
                Convert.ToDecimal(
                    value:
                        importItem.Height,
                    provider:
                        CultureInfo.InvariantCulture);

            int rowsInserted =
                await insertCoordinateCommand.ExecuteNonQueryAsync();

            if (rowsInserted != 1)
            {
                throw new InvalidOperationException(
                    $"Reference coordinates for '{importItem.SensorName}' were not inserted.");
            }

            #endregion
        }

        #endregion

        #endregion


        #region Prism Pair Import UI

        #region Prism Pair Preview Selection

        private void dgPrismPairImport_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Display Selected Prism Pair Validation

            if (dgPrismPairImport.SelectedItem
                is not PrismPairImportItem selectedItem)
            {
                return;
            }

            txtPrismPairImportStatus.Text =
                selectedItem.ValidationStatus;

            #endregion
        }

        #endregion


        #region Prism Pair Preview Extraction





        private void btnPreviewPrismPairs_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Refresh Prism Pair Preview

            PopulatePrismPairPreview();

            #endregion
        }


        private void PopulatePrismPairPreview()
        {
            #region Validate Preview Context

            if (string.IsNullOrWhiteSpace(_selectedPrismPairWorkbookPath))
            {
                dgPrismPairImport.ItemsSource =
                    null;

                txtPrismPairImportStatus.Text =
                    "Select a Track Geometry workbook.";

                return;
            }

            if (_prismPairImportConfiguration is null)
            {
                dgPrismPairImport.ItemsSource =
                    null;

                txtPrismPairImportStatus.Text =
                    "The Track Geometry import configuration is not loaded.";

                return;
            }

            if (_selectedPrismPairWorksheetProfile is null)
            {
                dgPrismPairImport.ItemsSource =
                    null;

                txtPrismPairImportStatus.Text =
                    "Select a supported worksheet.";

                return;
            }

            #endregion


            #region Resolve Column Roles

            if (!TryGetPrismPairColumnRoleMapping(
                rightRailSourceIndex: out int rightRailSourceIndex,
                leftRailSourceIndex: out int leftRailSourceIndex,
                railTagSourceIndex: out int railTagSourceIndex))
            {
                dgPrismPairImport.ItemsSource =
                    null;

                btnPreviewPrismPairs.IsEnabled =
                    false;

                btnCommitPrismPairImport.IsEnabled =
                    false;

                txtPrismPairImportStatus.Text =
                    "Right Rail, Left Rail and Rail Tag must each be assigned exactly once.";

                return;
            }

            #endregion


            try
            {
                #region Extract Prism Pairs

                List<PrismPairImportItem> prismPairs =
                    ExtractPrismPairPreview(
                        workbookPath: _selectedPrismPairWorkbookPath,
                        configuration: _prismPairImportConfiguration,
                        profile: _selectedPrismPairWorksheetProfile,
                        rightRailSourceIndex: rightRailSourceIndex,
                        leftRailSourceIndex: leftRailSourceIndex,
                        railTagSourceIndex: railTagSourceIndex,
                        railSectionCount: out int railSectionCount);

                #endregion


                #region Populate Prism Pair DataGrid

                _prismPairImportItems.Clear();

                foreach (PrismPairImportItem prismPair
                    in prismPairs)
                {
                    _prismPairImportItems.Add(
                        item: prismPair);
                }

                dgPrismPairImport.ItemsSource =
                    null;

                dgPrismPairImport.ItemsSource =
                    _prismPairImportItems;

                #endregion


                #region Report Preview Result

                btnCommitPrismPairImport.IsEnabled =
                    false;

                if (railSectionCount == 0)
                {
                    txtPrismPairImportStatus.Text =
                        $"No Rail Start marker was found in worksheet " +
                        $"'{_selectedPrismPairWorksheetProfile.WorksheetName}'.";

                    return;
                }

                if (_prismPairImportItems.Count == 0)
                {
                    txtPrismPairImportStatus.Text =
                        $"{railSectionCount} rail section(s) found, " +
                        "but no prism pairs were extracted.";

                    return;
                }


                int invalidRowCount =
                    0;

                foreach (PrismPairImportItem importItem
                    in _prismPairImportItems)
                {
                    if (!importItem.IsValid)
                    {
                        invalidRowCount++;
                    }
                }


                if (invalidRowCount > 0)
                {
                    txtPrismPairImportStatus.Text =
                        $"{_prismPairImportItems.Count} pair(s) extracted; " +
                        $"{invalidRowCount} row(s) contain validation errors.";

                    return;
                }


                if (!_prismPairImportProjectId.HasValue)
                {
                    txtPrismPairImportStatus.Text =
                        "Preview complete, but no target project is associated with the import.";

                    return;
                }


                if (!_activeProjectId.HasValue ||
                    _activeProjectId.Value != _prismPairImportProjectId.Value)
                {
                    txtPrismPairImportStatus.Text =
                        "Preview complete, but the active project has changed. " +
                        "Select the workbook again.";

                    return;
                }


                btnCommitPrismPairImport.IsEnabled =
                    true;

                txtPrismPairImportStatus.Text =
                    $"{_prismPairImportItems.Count} prism pair(s) extracted from " +
                    $"{railSectionCount} track(s). " +
                    $"Ready to commit to project '{_prismPairImportProjectName}'.";

                #endregion




            }
            catch (Exception ex)
            {
                #region Report Preview Failure

                dgPrismPairImport.ItemsSource =
                    null;

                btnCommitPrismPairImport.IsEnabled =
                    false;

                txtPrismPairImportStatus.Text =
                    $"Unable to extract prism pairs: {ex.Message}";

                #endregion
            }
        }


        private static List<PrismPairImportItem> ExtractPrismPairPreview(
            string workbookPath,
            TrackGeometryImportConfiguration configuration,
            TrackGeometryWorksheetProfile profile,
            int rightRailSourceIndex,
            int leftRailSourceIndex,
            int railTagSourceIndex,
            out int railSectionCount)
        {
            #region Validate Parameters

            if (string.IsNullOrWhiteSpace(workbookPath))
            {
                throw new ArgumentException(
                    message: "Workbook path cannot be blank.",
                    paramName: nameof(workbookPath));
            }

            ArgumentNullException.ThrowIfNull(
                argument: configuration);

            ArgumentNullException.ThrowIfNull(
                argument: profile);

            if (!File.Exists(
                path: workbookPath))
            {
                throw new FileNotFoundException(
                    message: "The selected Track Geometry workbook does not exist.",
                    fileName: workbookPath);
            }

            #endregion


            #region Open Workbook

            FileInfo workbookFile =
                new(
                    fileName: workbookPath);

            using ExcelPackage excelPackage =
                new(
                    newFile: workbookFile);

            #endregion


            #region Locate Worksheet

            ExcelWorksheet? selectedWorksheet =
                null;

            foreach (ExcelWorksheet worksheet
                in excelPackage.Workbook.Worksheets)
            {
                if (string.Equals(
                    a: worksheet.Name.Trim(),
                    b: profile.WorksheetName.Trim(),
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    selectedWorksheet =
                        worksheet;

                    break;
                }
            }

            if (selectedWorksheet is null)
            {
                throw new InvalidOperationException(
                    $"Worksheet '{profile.WorksheetName}' was not found.");
            }

            if (selectedWorksheet.Dimension is null)
            {
                throw new InvalidOperationException(
                    $"Worksheet '{profile.WorksheetName}' contains no used cells.");
            }

            #endregion


            #region Resolve Physical Source Columns

            int sourceColumn1 =
                ExcelColumnReferenceToNumber(
                    columnReference: profile.PrimaryRailColumn);

            int sourceColumn2 =
                ExcelColumnReferenceToNumber(
                    columnReference: profile.SecondaryRailColumn);

            int sourceColumn3 =
                ExcelColumnReferenceToNumber(
                    columnReference: profile.TrackLabelColumn);

            #endregion


            #region Initialise Extraction

            List<PrismPairImportItem> prismPairs =
                new();

            railSectionCount =
                0;

            int pairOrder =
                0;

            string currentTrackName =
                string.Empty;

            bool insideRailSection =
                false;

            int lastUsedRow =
                selectedWorksheet.Dimension.End.Row;

            #endregion


            #region Extract Worksheet Rail Sections

            for (int rowNumber = profile.StartRow;
                 rowNumber <= lastUsedRow;
                 rowNumber++)
            {
                #region Read Three Configured Source Values

                string sourceValue1 =
                    selectedWorksheet.Cells[
                        rowNumber,
                        sourceColumn1]
                    .Text
                    .Trim();

                string sourceValue2 =
                    selectedWorksheet.Cells[
                        rowNumber,
                        sourceColumn2]
                    .Text
                    .Trim();

                string sourceValue3 =
                    selectedWorksheet.Cells[
                        rowNumber,
                        sourceColumn3]
                    .Text
                    .Trim();

                #endregion


                #region Resolve Selected Roles

                string rightPointName =
                    GetPrismPairSourceValue(
                        sourceValue1: sourceValue1,
                        sourceValue2: sourceValue2,
                        sourceValue3: sourceValue3,
                        sourceIndex: rightRailSourceIndex);

                string leftPointName =
                    GetPrismPairSourceValue(
                        sourceValue1: sourceValue1,
                        sourceValue2: sourceValue2,
                        sourceValue3: sourceValue3,
                        sourceIndex: leftRailSourceIndex);

                string railTag =
                    GetPrismPairSourceValue(
                        sourceValue1: sourceValue1,
                        sourceValue2: sourceValue2,
                        sourceValue3: sourceValue3,
                        sourceIndex: railTagSourceIndex);

                #endregion


                #region Identify Rail Markers

                bool isRailStart =
                    IsConfiguredRailMarker(
                        value: railTag,
                        markers: configuration.StartMarkers);

                bool isRailEnd =
                    IsConfiguredRailMarker(
                        value: railTag,
                        markers: configuration.EndMarkers);

                #endregion


                #region Handle Rail Start

                if (isRailStart)
                {
                    railSectionCount++;

                    currentTrackName =
                        $"Track_{railSectionCount}";

                    pairOrder =
                        0;

                    insideRailSection =
                        true;
                }

                #endregion


                #region Ignore Rows Outside Rail Section

                if (!insideRailSection)
                {
                    continue;
                }

                #endregion


                #region Handle Blank Prism Pair

                if (string.IsNullOrWhiteSpace(leftPointName) &&
                    string.IsNullOrWhiteSpace(rightPointName))
                {
                    if (isRailStart)
                    {
                        continue;
                    }

                    insideRailSection =
                        false;

                    continue;
                }

                #endregion


                #region Determine Preview Validation

                string validationStatus;

                if (string.IsNullOrWhiteSpace(leftPointName))
                {
                    validationStatus =
                        "Left prism is blank.";
                }
                else if (string.IsNullOrWhiteSpace(rightPointName))
                {
                    validationStatus =
                        "Right prism is blank.";
                }
                else if (string.Equals(
                    a: leftPointName,
                    b: rightPointName,
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    validationStatus =
                        "Left and Right prism are identical.";
                }
                else
                {
                    validationStatus =
                        "Valid";
                }

                #endregion


                #region Add Prism Pair Preview Row

                pairOrder++;

                bool isValid =
                    string.Equals(
                        a: validationStatus,
                        b: "Valid",
                        comparisonType: StringComparison.Ordinal);

                PrismPairImportItem prismPair =
                    new()
                    {
                        SourceRow =
                            rowNumber,

                        RailSection =
                            railSectionCount,

                        TrackName =
                            currentTrackName,

                        PairOrder =
                            pairOrder,

                        SourceColumn1Value =
                            sourceValue1,

                        SourceColumn2Value =
                            sourceValue2,

                        SourceColumn3Value =
                            sourceValue3,

                        LeftPointName =
                            leftPointName,

                        RightPointName =
                            rightPointName,

                        IsValid =
                            isValid,

                        ValidationStatus =
                            validationStatus
                    };

                prismPairs.Add(
                    item: prismPair);

                #endregion


                #region Handle Rail End

                // The Rail End row is itself a valid prism-pair row.
                //
                // It must therefore be extracted before the current rail section
                // is closed.

                if (isRailEnd)
                {
                    insideRailSection =
                        false;
                }

                #endregion
            }

            #endregion


            #region Return Prism Pair Preview

            return prismPairs;

            #endregion
        }


        private static string GetPrismPairSourceValue(
            string sourceValue1,
            string sourceValue2,
            string sourceValue3,
            int sourceIndex)
        {
            #region Return Selected Source Value

            return sourceIndex switch
            {
                1 => sourceValue1,
                2 => sourceValue2,
                3 => sourceValue3,

                _ => throw new ArgumentOutOfRangeException(
                    paramName: nameof(sourceIndex),
                    message: "Prism Pair source index must be 1, 2 or 3.")
            };

            #endregion
        }


        #region Rail Marker Matching

        private static bool IsConfiguredRailMarker(
            string value,
            List<string> markers)
        {
            #region Validate Marker Value

            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            #endregion


            #region Normalise Source Marker

            string normalisedValue =
                NormaliseRailMarker(
                    value: value);

            #endregion


            #region Compare Configured Markers

            foreach (string marker in markers)
            {
                string normalisedMarker =
                    NormaliseRailMarker(
                        value: marker);

                if (string.Equals(
                    a: normalisedValue,
                    b: normalisedMarker,
                    comparisonType: StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;

            #endregion
        }


        private static string NormaliseRailMarker(
            string value)
        {
            #region Normalise Rail Marker

            return value
                .Trim()
                .Replace(
                    oldValue: " ",
                    newValue: string.Empty)
                .Replace(
                    oldValue: "_",
                    newValue: string.Empty)
                .ToUpperInvariant();

            #endregion
        }

        #endregion


        private static int ExcelColumnReferenceToNumber(
            string columnReference)
        {
            #region Validate Column Reference

            if (string.IsNullOrWhiteSpace(columnReference))
            {
                throw new ArgumentException(
                    message: "Excel column reference cannot be blank.",
                    paramName: nameof(columnReference));
            }

            #endregion


            #region Convert Excel Column Reference

            string validatedColumnReference =
                columnReference
                    .Trim()
                    .ToUpperInvariant();

            int columnNumber =
                0;

            foreach (char character in validatedColumnReference)
            {
                if (character < 'A' ||
                    character > 'Z')
                {
                    throw new ArgumentException(
                        message:
                            $"Invalid Excel column reference " +
                            $"'{columnReference}'.",
                        paramName: nameof(columnReference));
                }

                columnNumber =
                    (columnNumber * 26) +
                    (character - 'A' + 1);
            }

            return columnNumber;

            #endregion
        }

        #endregion

        #region Prism Pair Database Commit

        private async void btnCommitPrismPairImport_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Commit Context

            if (!_prismPairImportProjectId.HasValue)
            {
                btnCommitPrismPairImport.IsEnabled =
                    false;

                txtPrismPairImportStatus.Text =
                    "Import blocked: no project is associated with the preview.";

                return;
            }


            if (!_activeProjectId.HasValue ||
                _activeProjectId.Value != _prismPairImportProjectId.Value)
            {
                btnCommitPrismPairImport.IsEnabled =
                    false;

                txtPrismPairImportStatus.Text =
                    "Import blocked: the active project has changed. " +
                    "Select the workbook again.";

                return;
            }


            if (_prismPairImportItems.Count == 0)
            {
                btnCommitPrismPairImport.IsEnabled =
                    false;

                txtPrismPairImportStatus.Text =
                    "Import blocked: there are no prism pairs to import.";

                return;
            }


            foreach (PrismPairImportItem importItem
                in _prismPairImportItems)
            {
                if (!importItem.IsValid)
                {
                    btnCommitPrismPairImport.IsEnabled =
                        false;

                    txtPrismPairImportStatus.Text =
                        "Import blocked: the preview contains validation errors.";

                    return;
                }
            }

            #endregion


            #region Count Tracks

            HashSet<string> trackNames =
                new(
                    comparer: StringComparer.OrdinalIgnoreCase);

            foreach (PrismPairImportItem importItem
                in _prismPairImportItems)
            {
                trackNames.Add(
                    item: importItem.TrackName);
            }

            #endregion


            #region Confirm Prism Pair Import

            string confirmationMessage =
                $"Project: {_prismPairImportProjectName}\n\n" +
                $"Tracks to create: {trackNames.Count}\n" +
                $"Prism pairs to import: {_prismPairImportItems.Count}\n\n" +
                "Commit this import?";

            MessageBoxResult confirmationResult =
                MessageBox.Show(
                    owner: this,
                    messageBoxText: confirmationMessage,
                    caption: "Confirm Prism Pair Import",
                    button: MessageBoxButton.YesNo,
                    icon: MessageBoxImage.Question);

            if (confirmationResult != MessageBoxResult.Yes)
            {
                txtPrismPairImportStatus.Text =
                    "Import cancelled. No data was written to the database.";

                return;
            }

            #endregion


            #region Prepare Prism Pair Commit

            btnCommitPrismPairImport.IsEnabled =
                false;

            btnPreviewPrismPairs.IsEnabled =
                false;

            btnSelectPrismPairWorkbook.IsEnabled =
                false;

            btnClearPrismPairImport.IsEnabled =
                false;

            cmbPrismPairWorksheet.IsEnabled =
                false;

            cmbPrismPairColumn1Role.IsEnabled =
                false;

            cmbPrismPairColumn2Role.IsEnabled =
                false;

            cmbPrismPairColumn3Role.IsEnabled =
                false;

            txtPrismPairImportStatus.Text =
                $"Revalidating and importing {_prismPairImportItems.Count} " +
                $"prism pair(s) into project '{_prismPairImportProjectName}'...";

            #endregion


            try
            {
                #region Execute Atomic Prism Pair Import

                (int TrackCount, int PairCount) importResult =
                    await CommitPrismPairImportAsync(
                        projectId: _prismPairImportProjectId.Value,
                        projectName: _prismPairImportProjectName);

                #endregion


                #region Report Successful Prism Pair Import

                btnCommitPrismPairImport.IsEnabled =
                    false;

                txtPrismPairImportStatus.Text =
                    $"{importResult.TrackCount} track(s) and " +
                    $"{importResult.PairCount} prism pair(s) " +
                    $"imported successfully into project " +
                    $"'{_prismPairImportProjectName}'.";

                #endregion
            }
            catch (Exception ex)
            {
                #region Report Failed Prism Pair Import

                dgPrismPairImport.Items.Refresh();

                bool previewStillValid =
                    true;

                foreach (PrismPairImportItem importItem
                    in _prismPairImportItems)
                {
                    if (!importItem.IsValid)
                    {
                        previewStillValid =
                            false;

                        break;
                    }
                }

                btnCommitPrismPairImport.IsEnabled =
                    previewStillValid;

                txtPrismPairImportStatus.Text =
                    $"Import failed. No data committed. {ex.Message}";

                #endregion
            }
            finally
            {
                #region Restore Prism Pair Import Controls

                btnSelectPrismPairWorkbook.IsEnabled =
                    true;

                btnClearPrismPairImport.IsEnabled =
                    true;

                cmbPrismPairWorksheet.IsEnabled =
                    true;

                cmbPrismPairColumn1Role.IsEnabled =
                    true;

                cmbPrismPairColumn2Role.IsEnabled =
                    true;

                cmbPrismPairColumn3Role.IsEnabled =
                    true;

                btnPreviewPrismPairs.IsEnabled =
                    true;

                #endregion
            }
        }


        private async Task<(int TrackCount, int PairCount)>
            CommitPrismPairImportAsync(
                int projectId,
                string projectName)
        {
            #region Validate Prism Pair Commit Parameters

            if (projectId <= 0)
            {
                throw new ArgumentOutOfRangeException(
                    paramName: nameof(projectId),
                    message: "Project_ID must be greater than zero.");
            }

            if (string.IsNullOrWhiteSpace(projectName))
            {
                throw new ArgumentException(
                    message: "Project name cannot be blank.",
                    paramName: nameof(projectName));
            }

            #endregion


            #region Capture Prism Pair Preview Dataset

            List<PrismPairImportItem> importItems =
                new(
                    collection: _prismPairImportItems);

            if (importItems.Count == 0)
            {
                throw new InvalidOperationException(
                    "The Prism Pair preview contains no records.");
            }

            #endregion


            #region Resolve Prism Pair Database Connection String

            string databaseConnectionString =
                GetTrackGeometryConnectionString();

            #endregion




            #region Open Prism Pair Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionString);

            await databaseConnection.OpenAsync();

            #endregion


            #region Begin Serializable Prism Pair Transaction

            using SqlTransaction transaction =
                databaseConnection.BeginTransaction(
                    iso:
                        System.Data.IsolationLevel.Serializable);

            bool transactionCommitted =
                false;

            #endregion


            try
            {
                #region Acquire Prism Pair Import Transaction Lock

                await AcquirePrismPairImportTransactionLockAsync(
                    projectId: projectId,
                    databaseConnection: databaseConnection,
                    transaction: transaction);

                #endregion


                #region Revalidate Prism Pair Target Project

                await ValidateImportTargetProjectAsync(
                    projectId: projectId,
                    expectedProjectName: projectName,
                    databaseConnection: databaseConnection,
                    transaction: transaction);

                #endregion


                #region Load Prism Pair Point Names

                Dictionary<string, int> pointIds =
                    await LoadPrismPairPointIdsAsync(
                        projectId: projectId,
                        databaseConnection: databaseConnection,
                        transaction: transaction);

                #endregion


                #region Load Existing Track Names

                HashSet<string> existingTrackNames =
                    await LoadExistingTrackNamesAsync(
                        projectId: projectId,
                        databaseConnection: databaseConnection,
                        transaction: transaction);

                #endregion


                #region Load Existing Paired Points

                HashSet<int> existingPairedPointIds =
                    await LoadExistingPairedPointIdsAsync(
                        projectId: projectId,
                        databaseConnection: databaseConnection,
                        transaction: transaction);

                #endregion


                #region Final Prism Pair Validation

                ValidatePrismPairImportForCommit(
                    importItems: importItems,
                    pointIds: pointIds,
                    existingTrackNames: existingTrackNames,
                    existingPairedPointIds: existingPairedPointIds);

                dgPrismPairImport.Items.Refresh();


                int invalidRowCount =
                    0;

                foreach (PrismPairImportItem importItem
                    in importItems)
                {
                    if (!importItem.IsValid)
                    {
                        invalidRowCount++;
                    }
                }

                if (invalidRowCount > 0)
                {
                    throw new InvalidOperationException(
                        $"{invalidRowCount} row(s) failed final validation.");
                }

                #endregion


                #region Insert Tracks

                Dictionary<string, int> trackIds =
                    new(
                        comparer: StringComparer.OrdinalIgnoreCase);

                foreach (PrismPairImportItem importItem
                    in importItems)
                {
                    if (trackIds.ContainsKey(
                        key: importItem.TrackName))
                    {
                        continue;
                    }

                    int trackId =
                        await InsertTrackAsync(
                            projectId: projectId,
                            trackName: importItem.TrackName,
                            databaseConnection: databaseConnection,
                            transaction: transaction);

                    trackIds.Add(
                        key: importItem.TrackName,
                        value: trackId);
                }

                #endregion


                #region Insert Prism Pairs

                foreach (PrismPairImportItem importItem
                    in importItems)
                {
                    int trackId =
                        trackIds[importItem.TrackName];

                    int leftPointId =
                        pointIds[importItem.LeftPointName];

                    int rightPointId =
                        pointIds[importItem.RightPointName];

                    await InsertPrismPairAsync(
                        trackId: trackId,
                        pairOrder: importItem.PairOrder,
                        leftPointId: leftPointId,
                        rightPointId: rightPointId,
                        databaseConnection: databaseConnection,
                        transaction: transaction);
                }

                #endregion


                #region Commit Prism Pair Transaction

                transaction.Commit();

                transactionCommitted =
                    true;

                return
                (
                    TrackCount:
                        trackIds.Count,

                    PairCount:
                        importItems.Count
                );

                #endregion
            }
            catch
            {
                #region Roll Back Prism Pair Transaction

                if (!transactionCommitted)
                {
                    try
                    {
                        transaction.Rollback();
                    }
                    catch
                    {
                        // Preserve the original exception.
                    }
                }

                throw;

                #endregion
            }
        }


        private static async Task AcquirePrismPairImportTransactionLockAsync(
            int projectId,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Prism Pair Import Lock

            string lockResource =
                $"{PrismPairImportLockResourcePrefix}{projectId}";

            const string lockSql = """
        DECLARE @LockResult int;

        EXEC @LockResult = sys.sp_getapplock
            @Resource = @Resource,
            @LockMode = 'Exclusive',
            @LockOwner = 'Transaction',
            @LockTimeout = 0;

        SELECT @LockResult;
        """;

            #endregion


            #region Acquire Prism Pair Import Lock

            await using SqlCommand lockCommand =
                new(
                    cmdText: lockSql,
                    connection: databaseConnection,
                    transaction: transaction);

            lockCommand.Parameters.Add(
                parameterName: "@Resource",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 255)
                .Value =
                    lockResource;

            object? lockResultValue =
                await lockCommand.ExecuteScalarAsync();

            int lockResult =
                Convert.ToInt32(
                    value: lockResultValue,
                    provider: CultureInfo.InvariantCulture);

            if (lockResult < 0)
            {
                throw new InvalidOperationException(
                    "Another Prism Pair import is currently modifying this project.");
            }

            #endregion
        }


        #region Prism Pair Point Name Lookup

        private static async Task<Dictionary<string, int>>
            LoadPrismPairPointIdsAsync(
                int projectId,
                SqlConnection databaseConnection,
                SqlTransaction transaction)
        {
            #region Define Prism Pair Point Query

            // Track Geometry normally displays ReplacementName.
            //
            // Prism Pair import must therefore resolve each worksheet prism
            // against BOTH:
            //
            //     PointName
            //     ReplacementName
            //
            // Both names identify the same PointName_ID.

            const string pointSql = """
        SELECT
            PN.[PointName],
            PN.[ReplacementName],
            PN.[PointName_ID]
        FROM [dbo].[PointName] AS PN WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN [dbo].[CoordinatesReference] AS CR WITH (HOLDLOCK)
            ON CR.[PointName_ID] = PN.[PointName_ID]
                    AND CR.[IsDeleted] = 0
        WHERE
            PN.[Project_ID] = @Project_ID
            AND PN.[IsDeleted] = 0;
        """;

            #endregion


            #region Initialise Prism Pair Point Lookup

            Dictionary<string, int> pointIds =
                new(
                    comparer: StringComparer.OrdinalIgnoreCase);

            #endregion


            #region Execute Prism Pair Point Query

            await using SqlCommand pointCommand =
                new(
                    cmdText: pointSql,
                    connection: databaseConnection,
                    transaction: transaction);

            pointCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            await using SqlDataReader reader =
                await pointCommand.ExecuteReaderAsync();

            #endregion


            #region Register Point And Replacement Names

            while (await reader.ReadAsync())
            {
                string pointName =
                    reader.GetString(
                        i: 0)
                    .Trim();

                string replacementName =
                    reader.GetString(
                        i: 1)
                    .Trim();

                int pointId =
                    reader.GetInt32(
                        i: 2);


                #region Register Point Name

                RegisterPrismPairPointLookupName(
                    pointIds: pointIds,
                    lookupName: pointName,
                    pointId: pointId,
                    nameType: "PointName");

                #endregion


                #region Register Replacement Name

                if (!string.IsNullOrWhiteSpace(replacementName))
                {
                    RegisterPrismPairPointLookupName(
                        pointIds: pointIds,
                        lookupName: replacementName,
                        pointId: pointId,
                        nameType: "ReplacementName");
                }

                #endregion
            }

            #endregion


            #region Return Prism Pair Point Lookup

            return pointIds;

            #endregion
        }


        private static void RegisterPrismPairPointLookupName(
            Dictionary<string, int> pointIds,
            string lookupName,
            int pointId,
            string nameType)
        {
            #region Validate Lookup Name

            if (string.IsNullOrWhiteSpace(lookupName))
            {
                if (string.Equals(
                    a: nameType,
                    b: "PointName",
                    comparisonType: StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"PointName_ID {pointId} contains a blank PointName.");
                }

                return;
            }

            string validatedLookupName =
                lookupName.Trim();

            #endregion


            #region Check Existing Lookup

            if (pointIds.TryGetValue(
                key: validatedLookupName,
                value: out int existingPointId))
            {
                // PointName and ReplacementName may legitimately be identical
                // for the SAME database point.

                if (existingPointId == pointId)
                {
                    return;
                }

                // The same name resolving to two different PointName_ID values
                // would make the Track Geometry prism reference ambiguous.
                //
                // Never silently select one of them.

                throw new InvalidOperationException(
                    $"Database prism-name collision: '{validatedLookupName}' " +
                    $"resolves to both PointName_ID {existingPointId} and " +
                    $"PointName_ID {pointId}. " +
                    "PointName and ReplacementName values must form an " +
                    "unambiguous project namespace.");
            }

            #endregion


            #region Register Lookup Name

            pointIds.Add(
                key: validatedLookupName,
                value: pointId);

            #endregion
        }

        #endregion


        private static async Task<HashSet<string>>
            LoadExistingTrackNamesAsync(
                int projectId,
                SqlConnection databaseConnection,
                SqlTransaction transaction)
        {
            #region Define Existing Track Query

            const string trackSql = """
        SELECT
            [TrackName]
        FROM [dbo].[Track] WITH (UPDLOCK, HOLDLOCK)
        WHERE
            [Project_ID] = @Project_ID;
        """;

            #endregion


            #region Load Existing Track Names

            HashSet<string> trackNames =
                new(
                    comparer: StringComparer.OrdinalIgnoreCase);

            await using SqlCommand trackCommand =
                new(
                    cmdText: trackSql,
                    connection: databaseConnection,
                    transaction: transaction);

            trackCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            await using SqlDataReader reader =
                await trackCommand.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                trackNames.Add(
                    item:
                        reader.GetString(
                            i: 0));
            }

            return trackNames;

            #endregion
        }


        private static async Task<HashSet<int>>
            LoadExistingPairedPointIdsAsync(
                int projectId,
                SqlConnection databaseConnection,
                SqlTransaction transaction)
        {
            #region Define Existing Paired Point Query

            const string prismPairSql = """
        SELECT
            PP.[Left_ID],
            PP.[Right_ID]
        FROM [dbo].[PrismPairs] AS PP WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN [dbo].[Track] AS T WITH (HOLDLOCK)
            ON T.[Track_ID] = PP.[Track_ID]
        WHERE
            T.[Project_ID] = @Project_ID;
        """;

            #endregion


            #region Load Existing Paired Point IDs

            HashSet<int> pointIds =
                new();

            await using SqlCommand prismPairCommand =
                new(
                    cmdText: prismPairSql,
                    connection: databaseConnection,
                    transaction: transaction);

            prismPairCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            await using SqlDataReader reader =
                await prismPairCommand.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                pointIds.Add(
                    item:
                        reader.GetInt32(
                            i: 0));

                pointIds.Add(
                    item:
                        reader.GetInt32(
                            i: 1));
            }

            return pointIds;

            #endregion
        }


        private static void ValidatePrismPairImportForCommit(
            List<PrismPairImportItem> importItems,
            Dictionary<string, int> pointIds,
            HashSet<string> existingTrackNames,
            HashSet<int> existingPairedPointIds)
        {
            #region Initialise Final Prism Pair Validation

            HashSet<int> incomingPairedPointIds =
                new();

            HashSet<string> trackPairOrders =
                new(
                    comparer: StringComparer.OrdinalIgnoreCase);

            foreach (PrismPairImportItem importItem
                in importItems)
            {
                importItem.IsValid =
                    true;

                importItem.ValidationStatus =
                    "Valid";
            }

            #endregion


            #region Validate Individual Prism Pair Rows

            foreach (PrismPairImportItem importItem
                in importItems)
            {
                if (string.IsNullOrWhiteSpace(importItem.TrackName))
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message: "Track Tag is blank.");
                }

                if (importItem.PairOrder <= 0)
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message: "Pair Order must be greater than zero.");
                }


                if (existingTrackNames.Contains(
                    item: importItem.TrackName))
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message:
                            $"Track '{importItem.TrackName}' already exists in the project.");
                }


                string pairOrderKey =
                    $"{importItem.TrackName}\u001F{importItem.PairOrder}";

                if (!trackPairOrders.Add(
                    item: pairOrderKey))
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message:
                            $"Pair Order {importItem.PairOrder} is duplicated within " +
                            $"'{importItem.TrackName}'.");
                }


                if (!pointIds.TryGetValue(
                    key: importItem.LeftPointName,
                    value: out int leftPointId))
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message:
                            $"Left prism '{importItem.LeftPointName}': No PointName or ReplacementName match in this project");

                    continue;
                }


                if (!pointIds.TryGetValue(
                    key: importItem.RightPointName,
                    value: out int rightPointId))
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message:
                            $"Right prism '{importItem.RightPointName}': No PointName or ReplacementName match in this project");

                    continue;
                }


                if (leftPointId == rightPointId)
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message: "Left and Right prism are identical.");

                    continue;
                }


                if (existingPairedPointIds.Contains(
                    item: leftPointId))
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message:
                            $"Left prism '{importItem.LeftPointName}' is already used " +
                            "by an existing prism pair.");
                }


                if (existingPairedPointIds.Contains(
                    item: rightPointId))
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message:
                            $"Right prism '{importItem.RightPointName}' is already used " +
                            "by an existing prism pair.");
                }


                if (!incomingPairedPointIds.Add(
                    item: leftPointId))
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message:
                            $"Left prism '{importItem.LeftPointName}' is used more than once " +
                            "in this import.");
                }


                if (!incomingPairedPointIds.Add(
                    item: rightPointId))
                {
                    AddPrismPairValidationError(
                        importItem: importItem,
                        message:
                            $"Right prism '{importItem.RightPointName}' is used more than once " +
                            "in this import.");
                }
            }

            #endregion
        }


        private static void AddPrismPairValidationError(
            PrismPairImportItem importItem,
            string message)
        {
            #region Append Prism Pair Validation Error

            importItem.IsValid =
                false;

            if (string.IsNullOrWhiteSpace(importItem.ValidationStatus) ||
                string.Equals(
                    a: importItem.ValidationStatus,
                    b: "Valid",
                    comparisonType: StringComparison.Ordinal))
            {
                importItem.ValidationStatus =
                    message;
            }
            else
            {
                importItem.ValidationStatus +=
                    $" {message}";
            }

            #endregion
        }


        private static async Task<int> InsertTrackAsync(
            int projectId,
            string trackName,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Track Insert

            const string trackSql = """
        INSERT INTO [dbo].[Track]
        (
            [Project_ID],
            [TrackName],
            [IsDeleted]
        )
        VALUES
        (
            @Project_ID,
            @TrackName,
            0
        );

        SELECT CAST(SCOPE_IDENTITY() AS int);
        """;

            #endregion


            #region Insert Track

            await using SqlCommand trackCommand =
                new(
                    cmdText: trackSql,
                    connection: databaseConnection,
                    transaction: transaction);

            trackCommand.Parameters.Add(
                parameterName: "@Project_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    projectId;

            trackCommand.Parameters.Add(
                parameterName: "@TrackName",
                sqlDbType: System.Data.SqlDbType.NVarChar,
                size: 200)
                .Value =
                    trackName;

            object? trackIdValue =
                await trackCommand.ExecuteScalarAsync();

            if (trackIdValue is null ||
                trackIdValue == DBNull.Value)
            {
                throw new InvalidOperationException(
                    $"SQL Server did not return Track_ID for '{trackName}'.");
            }

            return Convert.ToInt32(
                value: trackIdValue,
                provider: CultureInfo.InvariantCulture);

            #endregion
        }


        private static async Task InsertPrismPairAsync(
            int trackId,
            int pairOrder,
            int leftPointId,
            int rightPointId,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Prism Pair Insert

            const string prismPairSql = """
        INSERT INTO [dbo].[PrismPairs]
        (
            [Track_ID],
            [PairOrder],
            [Left_ID],
            [Right_ID],
            [IsDeleted]
        )
        VALUES
        (
            @Track_ID,
            @PairOrder,
            @Left_ID,
            @Right_ID,
            0
        );
        """;

            #endregion


            #region Insert Prism Pair

            await using SqlCommand prismPairCommand =
                new(
                    cmdText: prismPairSql,
                    connection: databaseConnection,
                    transaction: transaction);

            prismPairCommand.Parameters.Add(
                parameterName: "@Track_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    trackId;

            prismPairCommand.Parameters.Add(
                parameterName: "@PairOrder",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    pairOrder;

            prismPairCommand.Parameters.Add(
                parameterName: "@Left_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    leftPointId;

            prismPairCommand.Parameters.Add(
                parameterName: "@Right_ID",
                sqlDbType: System.Data.SqlDbType.Int)
                .Value =
                    rightPointId;

            await prismPairCommand.ExecuteNonQueryAsync();

            #endregion
        }

        #endregion


        #region Prism Pair Profile Configuration

        private static TrackGeometryImportConfiguration
            LoadPrismPairImportConfiguration()
        {


            #region Locate Configuration File

            string configurationPath =
                IOPath.Combine(
                    path1: AppContext.BaseDirectory,
                    path2: PrismPairImportProfileFileName);


            // During normal deployment the JSON file must reside beside the executable.
            //
            // During Visual Studio development, also search upward from the executable
            // folder for the project copy of the configuration file.

            if (!File.Exists(
                path: configurationPath))
            {
                DirectoryInfo? searchDirectory =
                    new(
                        path: AppContext.BaseDirectory);

                while (searchDirectory is not null)
                {
                    string candidatePath =
                        IOPath.Combine(
                            path1: searchDirectory.FullName,
                            path2: PrismPairImportProfileFileName);

                    if (File.Exists(
                        path: candidatePath))
                    {
                        configurationPath =
                            candidatePath;

                        break;
                    }

                    searchDirectory =
                        searchDirectory.Parent;
                }
            }


            if (!File.Exists(
                path: configurationPath))
            {
                throw new FileNotFoundException(
                    message:
                        $"Prism Pair import configuration file " +
                        $"'{PrismPairImportProfileFileName}' was not found. " +
                        $"Application folder: '{AppContext.BaseDirectory}'.",
                    fileName: configurationPath);
            }

            #endregion


            #region Read Configuration File

            string configurationJson =
                File.ReadAllText(
                    path: configurationPath);

            if (string.IsNullOrWhiteSpace(configurationJson))
            {
                throw new InvalidOperationException(
                    $"Prism Pair import configuration file " +
                    $"'{PrismPairImportProfileFileName}' is empty.");
            }

            #endregion

            #region Deserialize Configuration

            TrackGeometryImportConfiguration configuration =
                JsonSerializer.Deserialize<TrackGeometryImportConfiguration>(
                    json: configurationJson,
                    options: PrismPairJsonSerializerOptions)
                ?? throw new InvalidOperationException(
                    "Prism Pair import configuration could not be read.");

            #endregion


            #region Validate Configuration

            ValidatePrismPairImportConfiguration(
                configuration: configuration);

            #endregion


            #region Return Configuration

            return configuration;

            #endregion
        }


        private static void ValidatePrismPairImportConfiguration(
            TrackGeometryImportConfiguration configuration)
        {
            #region Validate Configuration Object

            ArgumentNullException.ThrowIfNull(
                argument: configuration);

            #endregion


            #region Validate Rail Markers

            if (configuration.StartMarkers.Count == 0)
            {
                throw new InvalidOperationException(
                    "No rail start markers are defined.");
            }

            if (configuration.EndMarkers.Count == 0)
            {
                throw new InvalidOperationException(
                    "No rail end markers are defined.");
            }

            #endregion


            #region Validate Worksheet Profiles

            if (configuration.WorksheetProfiles.Count == 0)
            {
                throw new InvalidOperationException(
                    "No supported worksheet profiles are defined.");
            }

            HashSet<string> worksheetNames =
                new(
                    comparer: StringComparer.OrdinalIgnoreCase);

            foreach (TrackGeometryWorksheetProfile profile
                in configuration.WorksheetProfiles)
            {
                if (string.IsNullOrWhiteSpace(profile.WorksheetName))
                {
                    throw new InvalidOperationException(
                        "A worksheet profile contains a blank WorksheetName.");
                }

                if (!worksheetNames.Add(
                    item: profile.WorksheetName.Trim()))
                {
                    throw new InvalidOperationException(
                        $"Duplicate worksheet profile " +
                        $"'{profile.WorksheetName}' exists.");
                }

                if (profile.StartRow <= 0)
                {
                    throw new InvalidOperationException(
                        $"Worksheet profile '{profile.WorksheetName}' " +
                        "contains an invalid StartRow.");
                }

                if (!IsValidExcelColumnReference(
                    columnReference: profile.PrimaryRailColumn))
                {
                    throw new InvalidOperationException(
                        $"Worksheet profile '{profile.WorksheetName}' contains " +
                        $"invalid PrimaryRailColumn " +
                        $"'{profile.PrimaryRailColumn}'.");
                }

                if (!IsValidExcelColumnReference(
                    columnReference: profile.SecondaryRailColumn))
                {
                    throw new InvalidOperationException(
                        $"Worksheet profile '{profile.WorksheetName}' contains " +
                        $"invalid SecondaryRailColumn " +
                        $"'{profile.SecondaryRailColumn}'.");
                }

                if (!IsValidExcelColumnReference(
                    columnReference: profile.TrackLabelColumn))
                {
                    throw new InvalidOperationException(
                        $"Worksheet profile '{profile.WorksheetName}' contains " +
                        $"invalid TrackLabelColumn " +
                        $"'{profile.TrackLabelColumn}'.");
                }
            }

            #endregion
        }


        private static bool IsValidExcelColumnReference(
            string columnReference)
        {
            #region Validate Column Reference

            if (string.IsNullOrWhiteSpace(columnReference))
            {
                return false;
            }

            string validatedColumnReference =
                columnReference.Trim();

            foreach (char character in validatedColumnReference)
            {
                if (!char.IsLetter(
                    c: character))
                {
                    return false;
                }
            }

            return true;

            #endregion
        }


        private static TrackGeometryWorksheetProfile?
            FindPrismPairWorksheetProfile(
                TrackGeometryImportConfiguration configuration,
                string worksheetName)
        {
            #region Locate Worksheet Profile

            foreach (TrackGeometryWorksheetProfile profile
                in configuration.WorksheetProfiles)
            {
                if (string.Equals(
                    a: profile.WorksheetName,
                    b: worksheetName,
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    return profile;
                }
            }

            return null;

            #endregion
        }

        #endregion


        #region Select Track Geometry Workbook

        private void btnSelectPrismPairWorkbook_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Select Workbook

            OpenFileDialog openFileDialog =
                new()
                {
                    Title =
                        "Select Track Geometry Workbook",

                    Filter =
                        "Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|" +
                        "Excel workbook (*.xlsx)|*.xlsx|" +
                        "Excel macro-enabled workbook (*.xlsm)|*.xlsm|" +
                        "All files (*.*)|*.*",

                    DefaultExt =
                        ".xlsx",

                    CheckFileExists =
                        true,

                    Multiselect =
                        false
                };

            bool? fileSelected =
                openFileDialog.ShowDialog(
                    owner: this);

            if (fileSelected != true)
            {
                txtPrismPairImportStatus.Text =
                    "Workbook selection cancelled.";

                return;
            }

            #endregion


            #region Validate Active Project

            if (!_activeProjectId.HasValue)
            {
                txtPrismPairImportStatus.Text =
                    "Workbook selected, but no active project is available.";

                MessageBox.Show(
                    owner: this,
                    messageBoxText:
                        "No active project is currently available.\n\n" +
                        "Select an active project using Manage Projects, " +
                        "then select the workbook again.",
                    caption: "Active Project Required",
                    button: MessageBoxButton.OK,
                    icon: MessageBoxImage.Warning);

                return;
            }

            if (string.IsNullOrWhiteSpace(
                value: _activeProjectName))
            {
                txtPrismPairImportStatus.Text =
                    "Workbook selected, but the active project name is unavailable.";

                MessageBox.Show(
                    owner: this,
                    messageBoxText:
                        "The active project does not contain a valid project name.\n\n" +
                        "Select the project again using Manage Projects.",
                    caption: "Invalid Active Project",
                    button: MessageBoxButton.OK,
                    icon: MessageBoxImage.Warning);

                return;
            }

            #endregion


            #region Capture Prism Pair Import Project Context

            int importProjectId =
                _activeProjectId.Value;

            string importProjectName =
                _activeProjectName;

            _prismPairImportProjectId =
                importProjectId;

            _prismPairImportProjectName =
                importProjectName;

            #endregion


            #region Store Selected Workbook

            _selectedPrismPairWorkbookPath =
                openFileDialog.FileName;

            txtPrismPairWorkbookPath.Text =
                _selectedPrismPairWorkbookPath;

            #endregion


            #region Reset Previous Workbook State

            _prismPairImportConfiguration =
                null;

            _selectedPrismPairWorksheetProfile =
                null;

            _prismPairImportItems.Clear();

            cmbPrismPairWorksheet.Items.Clear();

            dgPrismPairImport.ItemsSource =
                null;

            btnPreviewPrismPairs.IsEnabled =
                false;

            btnCommitPrismPairImport.IsEnabled =
                false;

            txtPrismPairImportStatus.Text =
                $"Reading supported worksheets from " +
                $"'{IOPath.GetFileName(_selectedPrismPairWorkbookPath)}'...";

            #endregion


            #region Load Supported Worksheets

            try
            {
                LoadPrismPairWorksheetNames(
                    workbookPath:
                        _selectedPrismPairWorkbookPath);
            }
            catch (Exception ex)
            {
                #region Reset Failed Workbook Load

                _prismPairImportConfiguration =
                    null;

                _selectedPrismPairWorksheetProfile =
                    null;

                _prismPairImportItems.Clear();

                cmbPrismPairWorksheet.Items.Clear();

                dgPrismPairImport.ItemsSource =
                    null;

                btnPreviewPrismPairs.IsEnabled =
                    false;

                btnCommitPrismPairImport.IsEnabled =
                    false;

                #endregion


                #region Report Workbook Failure

                txtPrismPairImportStatus.Text =
                    $"Unable to read workbook configuration: {ex.Message}";

                MessageBox.Show(
                    owner: this,
                    messageBoxText:
                        $"The selected workbook could not be processed.\n\n" +
                        $"{ex.Message}",
                    caption: "Workbook Import Error",
                    button: MessageBoxButton.OK,
                    icon: MessageBoxImage.Error);

                #endregion
            }

            #endregion
        }

        #endregion


        #region Clear Prism Pair Import

        private void btnClearPrismPairImport_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Reset Prism Pair Import

            ResetPrismPairImportState(
                statusMessage: "No workbook selected.");

            #endregion
        }


        private void ResetPrismPairImportState(
            string statusMessage)
        {
            #region Clear Prism Pair Import Project Context

            _prismPairImportProjectId =
                null;

            _prismPairImportProjectName =
                string.Empty;

            #endregion


            #region Clear Prism Pair Import Source State

            _selectedPrismPairWorkbookPath =
                string.Empty;

            _prismPairImportConfiguration =
                null;

            _selectedPrismPairWorksheetProfile =
                null;

            _prismPairImportItems.Clear();

            #endregion


            #region Clear Prism Pair User Interface

            txtPrismPairWorkbookPath.Text =
                string.Empty;

            cmbPrismPairWorksheet.Items.Clear();

            dgPrismPairImport.ItemsSource =
                null;

            btnPreviewPrismPairs.IsEnabled =
                false;

            btnCommitPrismPairImport.IsEnabled =
                false;

            txtPrismPairImportStatus.Text =
                statusMessage;

            #endregion
        }

        #endregion


        #region Load Prism Pair Worksheet Names

        private void LoadPrismPairWorksheetNames(
            string workbookPath)
        {
            #region Validate Workbook

            if (string.IsNullOrWhiteSpace(workbookPath))
            {
                throw new ArgumentException(
                    message: "Workbook path cannot be blank.",
                    paramName: nameof(workbookPath));
            }

            if (!File.Exists(
                path: workbookPath))
            {
                throw new FileNotFoundException(
                    message: "The selected Track Geometry workbook does not exist.",
                    fileName: workbookPath);
            }

            #endregion


            #region Load Import Configuration

            _prismPairImportConfiguration =
                LoadPrismPairImportConfiguration();

            _selectedPrismPairWorksheetProfile =
                null;

            #endregion


            #region Read Workbook Worksheet Names

            HashSet<string> workbookWorksheetNames =
                new(
                    comparer: StringComparer.OrdinalIgnoreCase);

            FileInfo workbookFile =
                new(
                    fileName: workbookPath);

            using (ExcelPackage excelPackage =
                new(
                    newFile: workbookFile))
            {
                foreach (ExcelWorksheet worksheet
                    in excelPackage.Workbook.Worksheets)
                {
                    string worksheetName =
                        worksheet.Name.Trim();

                    if (!string.IsNullOrWhiteSpace(worksheetName))
                    {
                        workbookWorksheetNames.Add(
                            item: worksheetName);
                    }
                }
            }

            // IMPORTANT:
            //
            // The ExcelPackage has now been disposed and the workbook released
            // before cmbPrismPairWorksheet.SelectedIndex is changed.
            //
            // Changing SelectedIndex raises cmbPrismPairWorksheet_SelectionChanged(),
            // which subsequently opens the workbook again to generate the preview.

            #endregion


            #region Populate Supported Worksheet List

            cmbPrismPairWorksheet.Items.Clear();

            foreach (TrackGeometryWorksheetProfile profile
                in _prismPairImportConfiguration.WorksheetProfiles)
            {
                string configuredWorksheetName =
                    profile.WorksheetName.Trim();

                if (workbookWorksheetNames.Contains(
                    item: configuredWorksheetName))
                {
                    cmbPrismPairWorksheet.Items.Add(
                        newItem: configuredWorksheetName);
                }
            }

            #endregion


            #region Validate Supported Worksheet List

            if (cmbPrismPairWorksheet.Items.Count == 0)
            {
                string workbookNames =
                    workbookWorksheetNames.Count > 0
                        ? string.Join(
                            separator: ", ",
                            values: workbookWorksheetNames)
                        : "<none>";

                string configuredNames =
                    string.Join(
                        separator: ", ",
                        values:
                            _prismPairImportConfiguration
                                .WorksheetProfiles
                                .ConvertAll(
                                    converter:
                                        profile =>
                                            profile.WorksheetName));

                txtPrismPairImportStatus.Text =
                    $"No supported worksheets found. " +
                    $"Workbook worksheets: [{workbookNames}] " +
                    $"Configured worksheets: [{configuredNames}]";

                return;
            }

            #endregion


            #region Select First Supported Worksheet

            // The workbook has already been released above.
            //
            // This selection raises cmbPrismPairWorksheet_SelectionChanged(),
            // which may safely reopen the workbook for Prism Pair extraction.

            cmbPrismPairWorksheet.SelectedIndex =
                0;

            #endregion
        }

        #endregion


        #region Prism Pair Worksheet Selection

        private void cmbPrismPairWorksheet_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Validate Selection State

            if (string.IsNullOrWhiteSpace(_selectedPrismPairWorkbookPath))
            {
                return;
            }

            if (_prismPairImportConfiguration is null)
            {
                return;
            }

            if (cmbPrismPairWorksheet.SelectedItem
                is not string worksheetName ||
                string.IsNullOrWhiteSpace(worksheetName))
            {
                return;
            }

            #endregion


            #region Validate Active Project

            if (!_activeProjectId.HasValue)
            {
                _prismPairImportProjectId =
                    null;

                _prismPairImportProjectName =
                    string.Empty;

                dgPrismPairImport.ItemsSource =
                    null;

                btnPreviewPrismPairs.IsEnabled =
                    false;

                btnCommitPrismPairImport.IsEnabled =
                    false;

                txtPrismPairImportStatus.Text =
                    "Select an active project before importing prism pairs.";

                return;
            }

            if (string.IsNullOrWhiteSpace(_activeProjectName))
            {
                _prismPairImportProjectId =
                    null;

                _prismPairImportProjectName =
                    string.Empty;

                dgPrismPairImport.ItemsSource =
                    null;

                btnPreviewPrismPairs.IsEnabled =
                    false;

                btnCommitPrismPairImport.IsEnabled =
                    false;

                txtPrismPairImportStatus.Text =
                    "The active project does not contain a valid project name.";

                return;
            }

            #endregion


            #region Locate Worksheet Profile

            TrackGeometryWorksheetProfile? selectedProfile =
                FindPrismPairWorksheetProfile(
                    configuration: _prismPairImportConfiguration,
                    worksheetName: worksheetName);

            if (selectedProfile is null)
            {
                _selectedPrismPairWorksheetProfile =
                    null;

                dgPrismPairImport.ItemsSource =
                    null;

                btnPreviewPrismPairs.IsEnabled =
                    false;

                btnCommitPrismPairImport.IsEnabled =
                    false;

                txtPrismPairImportStatus.Text =
                    $"No import profile exists for worksheet '{worksheetName}'.";

                return;
            }

            _selectedPrismPairWorksheetProfile =
                selectedProfile;

            #endregion


            #region Capture Prism Pair Import Project Context

            // Capture the process-local active project at the point at which
            // the worksheet preview is defined.
            //
            // Registry values are not used as operational import state.

            _prismPairImportProjectId =
                _activeProjectId.Value;

            _prismPairImportProjectName =
                _activeProjectName;

            #endregion


            #region Reset Previous Preview

            dgPrismPairImport.ItemsSource =
                null;

            btnCommitPrismPairImport.IsEnabled =
                false;

            #endregion


            #region Apply Default Column Roles

            SetDefaultPrismPairColumnRoles();

            #endregion


            #region Enable Preview

            btnPreviewPrismPairs.IsEnabled =
                true;

            txtPrismPairImportStatus.Text =
                $"Worksheet '{worksheetName}' selected. " +
                $"Source columns: " +
                $"{selectedProfile.PrimaryRailColumn}, " +
                $"{selectedProfile.SecondaryRailColumn}, " +
                $"{selectedProfile.TrackLabelColumn}.";

            #endregion


            #region Populate Preview

            PopulatePrismPairPreview();

            #endregion
        }

        #endregion


        #region Prism Pair Column Role Selection

        private void SetDefaultPrismPairColumnRoles()
        {
            #region Apply Profile Default Roles

            _isUpdatingPrismPairColumnRoles =
                true;

            try
            {
                cmbPrismPairColumn1Role.SelectedIndex =
                    0;

                cmbPrismPairColumn2Role.SelectedIndex =
                    1;

                cmbPrismPairColumn3Role.SelectedIndex =
                    2;
            }
            finally
            {
                _isUpdatingPrismPairColumnRoles =
                    false;
            }

            #endregion
        }


        private void PrismPairColumnRole_SelectionChanged(
            object sender,
            SelectionChangedEventArgs e)
        {
            #region Ignore Internal Role Updates

            if (_isUpdatingPrismPairColumnRoles)
            {
                return;
            }

            if (!IsInitialized)
            {
                return;
            }

            #endregion


            #region Validate Column Role Allocation

            if (!TryGetPrismPairColumnRoleMapping(
                rightRailSourceIndex: out _,
                leftRailSourceIndex: out _,
                railTagSourceIndex: out _))
            {
                dgPrismPairImport.ItemsSource =
                    null;

                btnPreviewPrismPairs.IsEnabled =
                    false;

                btnCommitPrismPairImport.IsEnabled =
                    false;

                txtPrismPairImportStatus.Text =
                    "Right Rail, Left Rail and Rail Tag must each be assigned exactly once.";

                return;
            }

            #endregion


            #region Refresh Preview

            btnPreviewPrismPairs.IsEnabled =
                true;

            btnCommitPrismPairImport.IsEnabled =
                false;

            if (!string.IsNullOrWhiteSpace(_selectedPrismPairWorkbookPath) &&
                _prismPairImportConfiguration is not null &&
                _selectedPrismPairWorksheetProfile is not null)
            {
                PopulatePrismPairPreview();
            }

            #endregion
        }


        private bool TryGetPrismPairColumnRoleMapping(
            out int rightRailSourceIndex,
            out int leftRailSourceIndex,
            out int railTagSourceIndex)
        {
            #region Initialise Role Mapping

            rightRailSourceIndex =
                0;

            leftRailSourceIndex =
                0;

            railTagSourceIndex =
                0;

            #endregion


            #region Read Selected Roles

            string role1 =
                GetSelectedPrismPairColumnRole(
                    comboBox: cmbPrismPairColumn1Role);

            string role2 =
                GetSelectedPrismPairColumnRole(
                    comboBox: cmbPrismPairColumn2Role);

            string role3 =
                GetSelectedPrismPairColumnRole(
                    comboBox: cmbPrismPairColumn3Role);

            if (string.IsNullOrWhiteSpace(role1) ||
                string.IsNullOrWhiteSpace(role2) ||
                string.IsNullOrWhiteSpace(role3))
            {
                return false;
            }

            #endregion


            #region Reject Duplicate Roles

            HashSet<string> selectedRoles =
                new(
                    comparer: StringComparer.OrdinalIgnoreCase)
                {
            role1,
            role2,
            role3
                };

            if (selectedRoles.Count != 3)
            {
                return false;
            }

            #endregion


            #region Resolve Source Index For Each Role

            string[] roles =
            {
        role1,
        role2,
        role3
    };

            for (int roleIndex = 0;
                 roleIndex < roles.Length;
                 roleIndex++)
            {
                int sourceIndex =
                    roleIndex + 1;

                if (string.Equals(
                    a: roles[roleIndex],
                    b: PrismPairRoleRightRail,
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    rightRailSourceIndex =
                        sourceIndex;
                }
                else if (string.Equals(
                    a: roles[roleIndex],
                    b: PrismPairRoleLeftRail,
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    leftRailSourceIndex =
                        sourceIndex;
                }
                else if (string.Equals(
                    a: roles[roleIndex],
                    b: PrismPairRoleRailTag,
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    railTagSourceIndex =
                        sourceIndex;
                }
                else
                {
                    return false;
                }
            }

            #endregion


            #region Validate Complete Mapping

            return
                rightRailSourceIndex > 0 &&
                leftRailSourceIndex > 0 &&
                railTagSourceIndex > 0;

            #endregion
        }


        private static string GetSelectedPrismPairColumnRole(
            ComboBox comboBox)
        {
            #region Read ComboBox Role

            if (comboBox.SelectedItem
                is not ComboBoxItem selectedItem)
            {
                return string.Empty;
            }

            return selectedItem.Content?.ToString()?.Trim()
                ?? string.Empty;

            #endregion
        }

        #endregion






        #endregion
    }

    #region Prism Pair Import Item

    public sealed class PrismPairImportItem
    {
        #region Worksheet Source

        public int SourceRow { get; init; }

        public int RailSection { get; init; }

        public int PairOrder { get; init; }

        public string TrackName { get; init; } =
            string.Empty;

        public string SourceColumn1Value { get; init; } =
            string.Empty;

        public string SourceColumn2Value { get; init; } =
            string.Empty;

        public string SourceColumn3Value { get; init; } =
            string.Empty;

        #endregion


        #region Pair Definition

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
