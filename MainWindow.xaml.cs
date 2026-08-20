#region System Preparation

using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows;


#region Import Processing Namespaces

using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using OfficeOpenXml;

#endregion

#region GNA classes
using GNAgeneraltools;

#endregion


#region WPF Control Namespaces

using System.Windows.Controls;

#endregion


using Microsoft.Data.SqlClient;
using Microsoft.Win32;

using GNAspreadsheettools;

#region Disable warnings

#pragma warning disable IDE0028
#pragma warning disable IDE0031
#pragma warning disable IDE0042
#pragma warning disable IDE0059
#pragma warning disable IDE0079
#pragma warning disable IDE0300

#pragma warning disable IDE1006
#pragma warning disable IDE0306

#endregion

#endregion

namespace GNA_DLRreport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {


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

        #endregion


        #region Project Configuration State

        // Project records currently displayed in the Project Management DataGrid.
        private readonly ObservableCollection<ProjectConfigurationItem> _projectItems =
            new();


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
        // Project deletion will later require an Exclusive lock on the same
        // resource. The Exclusive lock cannot be obtained while any running
        // instance holds a Shared lock.
        //
        // The dedicated lock connection uses Pooling = false so disposing the
        // connection terminates its SQL session and releases its Session-owned
        // application lock.
        // ---------------------------------------------------------------------

        private SqlConnection? _activeProjectLockConnection;

        private int? _activeProjectLockProjectId;

        private const string ProjectLockResourcePrefix =
            "GNA_DLRreport:Project:";

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
            0.0025m;

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

        #region Reference Coordinate Import

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


            #region Read Prism Pair Database Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "Database connection string has not been configured.");
            }

            SqlConnectionStringBuilder databaseConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog =
                        TrackGeometryDatabaseName
                };

            #endregion


            #region Ensure Track And Prism Pair Database Schema

            await EnsureTrackAndPrismPairSchemaAsync(
                connectionString:
                    databaseConnectionBuilder.ConnectionString);

            #endregion




            #region Open Prism Pair Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionBuilder.ConnectionString);

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

                await ValidatePrismPairImportTargetProjectAsync(
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


        private static async Task ValidatePrismPairImportTargetProjectAsync(
            int projectId,
            string expectedProjectName,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Prism Pair Project Validation Query

            const string projectSql = """
        SELECT
            [ProjectName]
        FROM [dbo].[Project] WITH (UPDLOCK, HOLDLOCK)
        WHERE
            [Project_ID] = @Project_ID
            AND [IsDeleted] = 0;
        """;

            #endregion


            #region Read Prism Pair Target Project

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
                    "The target project no longer exists or has been deleted.");
            }

            string databaseProjectName =
                Convert.ToString(
                    value: projectNameValue,
                    provider: CultureInfo.InvariantCulture)
                ?? throw new InvalidOperationException(
                    "The target project name could not be read.");

            #endregion


            #region Verify Prism Pair Project Name

            if (!string.Equals(
                a: databaseProjectName,
                b: expectedProjectName,
                comparisonType: StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    $"The target project name has changed from " +
                    $"'{expectedProjectName}' to '{databaseProjectName}'. " +
                    "Select the workbook again.");
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
                Path.Combine(
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
                        Path.Combine(
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
                $"'{Path.GetFileName(_selectedPrismPairWorkbookPath)}'...";

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

                List<ExistingReferencePoint> existingPoints =
                    await LoadExistingReferencePointsAsync(
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

                // Block 3 will attach the actual database commit operation.
                //
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

        private async Task<List<ExistingReferencePoint>>
            LoadExistingReferencePointsAsync(
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


            #region Read Database Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "Database connection string has not been configured.");
            }

            #endregion


            #region Build Database Connection String

            SqlConnectionStringBuilder databaseConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = TrackGeometryDatabaseName
                };

            #endregion


            #region Define Existing Point Query

            // Deliberately include soft-deleted PointName records.
            //
            // They still exist within the project namespace and their PointName
            // remains subject to the database uniqueness constraint.

            const string existingPointsSql = """
        SELECT
            PN.[PointName_ID],
            PN.[Project_ID],
            PN.[PointName],
            PN.[ReplacementName],
            CR.[Eref],
            CR.[Nref],
            CR.[Href]
        FROM [dbo].[PointName] AS PN
        INNER JOIN [dbo].[CoordinatesReference] AS CR
            ON CR.[PointName_ID] = PN.[PointName_ID]
        WHERE
            PN.[Project_ID] = @Project_ID;
        """;

            #endregion


            #region Load Existing Points

            List<ExistingReferencePoint> existingPoints =
                new();

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionBuilder.ConnectionString);

            await databaseConnection.OpenAsync();

            await using SqlCommand existingPointsCommand =
                new(
                    cmdText: existingPointsSql,
                    connection: databaseConnection);

            existingPointsCommand.Parameters.AddWithValue(
                parameterName: "@Project_ID",
                value: projectId);

            await using SqlDataReader reader =
                await existingPointsCommand.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                ExistingReferencePoint existingPoint =
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

                        Easting =
                            reader.GetDecimal(
                                i: 4),

                        Northing =
                            reader.GetDecimal(
                                i: 5),

                        Height =
                            reader.GetDecimal(
                                i: 6)
                    };

                existingPoints.Add(
                    item: existingPoint);
            }

            #endregion


            #region Return Existing Points

            return existingPoints;

            #endregion
        }

        #endregion


        #region Complete Import Validation

        private static void ValidateReferenceCoordinateImport(
            List<ReferenceCoordinateImportItem> importItems,
            List<ExistingReferencePoint> existingPoints)
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

                // PointName == ReplacementName on the SAME record is permitted.
                // This is the normal condition when no ReplacementName was supplied.

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


            #region Build Existing Combined Name Namespace

            Dictionary<string,
                (ExistingReferencePoint Point, string NameType)>
                existingNames =
                    new(
                        comparer: StringComparer.OrdinalIgnoreCase);

            foreach (ExistingReferencePoint existingPoint in existingPoints)
            {
                RegisterExistingName(
                    name: existingPoint.PointName,
                    nameType: "Point Name",
                    existingPoint: existingPoint,
                    existingNames: existingNames);

                if (!string.Equals(
                    a: existingPoint.PointName,
                    b: existingPoint.ReplacementName,
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    RegisterExistingName(
                        name: existingPoint.ReplacementName,
                        nameType: "Replacement Name",
                        existingPoint: existingPoint,
                        existingNames: existingNames);
                }
            }

            #endregion


            #region Validate Incoming Names Against Database

            foreach (ReferenceCoordinateImportItem importItem in importItems)
            {
                ValidateNameAgainstExistingDatabase(
                    name: importItem.PointName,
                    nameType: "Point Name",
                    importItem: importItem,
                    existingNames: existingNames);

                if (!string.Equals(
                    a: importItem.PointName,
                    b: importItem.ReplacementName,
                    comparisonType: StringComparison.OrdinalIgnoreCase))
                {
                    ValidateNameAgainstExistingDatabase(
                        name: importItem.ReplacementName,
                        nameType: "Replacement Name",
                        importItem: importItem,
                        existingNames: existingNames);
                }
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
                        $"{nameType} '{name}' duplicates " +
                        $"{existingName.NameType} on CSV row " +
                        $"{existingName.Item.SourceRow}.");

                AppendValidationError(
                    importItem: existingName.Item,
                    errorMessage:
                        $"{existingName.NameType} '{name}' duplicates " +
                        $"{nameType} on CSV row {importItem.SourceRow}.");

                return;
            }

            #endregion


            #region Register Name

            incomingNames.Add(
                key: name,
                value: (importItem, nameType));

            #endregion
        }


        private static void RegisterExistingName(
            string name,
            string nameType,
            ExistingReferencePoint existingPoint,
            Dictionary<string,
                (ExistingReferencePoint Point, string NameType)> existingNames)
        {
            #region Ignore Blank Existing Names

            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            #endregion


            #region Register Existing Name

            // If inconsistent historic database data already contains a combined
            // namespace collision, retaining the first record is sufficient for
            // import validation. Any incoming use of the same name is blocked.

            if (!existingNames.ContainsKey(
                key: name))
            {
                existingNames.Add(
                    key: name,
                    value: (existingPoint, nameType));
            }

            #endregion
        }


        private static void ValidateNameAgainstExistingDatabase(
            string name,
            string nameType,
            ReferenceCoordinateImportItem importItem,
            Dictionary<string,
                (ExistingReferencePoint Point, string NameType)> existingNames)
        {
            #region Ignore Invalid Blank Names

            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            #endregion


            #region Check Existing Database Namespace

            if (existingNames.TryGetValue(
                key: name,
                value: out
                    (ExistingReferencePoint Point, string NameType)
                    existingName))
            {
                AppendValidationError(
                    importItem: importItem,
                    errorMessage:
                        $"{nameType} '{name}' duplicates existing " +
                        $"{existingName.NameType} on database point " +
                        $"'{existingName.Point.PointName}'.");
            }

            #endregion
        }

        #endregion


        #region Horizontal Coordinate Duplicate Validation

        private static void ValidateHorizontalCoordinateDuplicates(
            List<ReferenceCoordinateImportItem> importItems,
            List<ExistingReferencePoint> existingPoints)
        {
            #region Build Existing Coordinate Buckets

            Dictionary<(long EastingBucket, long NorthingBucket),
                List<ExistingReferencePoint>>
                existingCoordinateBuckets =
                    new();

            foreach (ExistingReferencePoint existingPoint in existingPoints)
            {
                (long EastingBucket, long NorthingBucket) bucket =
                    GetReferenceCoordinateBucket(
                        easting: existingPoint.Easting,
                        northing: existingPoint.Northing);

                if (!existingCoordinateBuckets.TryGetValue(
                    key: bucket,
                    value: out List<ExistingReferencePoint>? bucketPoints))
                {
                    bucketPoints =
                        new List<ExistingReferencePoint>();

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
                                List<ExistingReferencePoint>? existingBucketPoints))
                        {
                            continue;
                        }

                        foreach (ExistingReferencePoint existingPoint
                            in existingBucketPoints)
                        {
                            decimal distanceSquared =
                                GetHorizontalDistanceSquared(
                                    easting1: easting,
                                    northing1: northing,
                                    easting2: existingPoint.Easting,
                                    northing2: existingPoint.Northing);

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


            #region Read Database Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "Database connection string has not been configured.");
            }

            SqlConnectionStringBuilder databaseConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = TrackGeometryDatabaseName
                };

            #endregion


            #region Open Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionBuilder.ConnectionString);

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

                await ValidateReferenceImportTargetProjectAsync(
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

                List<ExistingReferencePoint> existingPoints =
                    await LoadExistingReferencePointsForCommitAsync(
                        projectId: projectId,
                        databaseConnection: databaseConnection,
                        transaction: transaction);

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
                        $"{invalidRowCount} row(s) failed final validation. " +
                        "The database may have changed since the preview was created.");
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


        #region Reference Import Project Revalidation

        private static async Task ValidateReferenceImportTargetProjectAsync(
            int projectId,
            string expectedProjectName,
            SqlConnection databaseConnection,
            SqlTransaction transaction)
        {
            #region Define Project Validation Query

            const string projectSql = """
        SELECT
            [ProjectName]
        FROM [dbo].[Project] WITH (UPDLOCK, HOLDLOCK)
        WHERE
            [Project_ID] = @Project_ID
            AND [IsDeleted] = 0;
        """;

            #endregion


            #region Read Current Project

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
                    "The target project no longer exists or has been deleted.");
            }

            string databaseProjectName =
                Convert.ToString(
                    value: projectNameValue,
                    provider: CultureInfo.InvariantCulture)
                ?? throw new InvalidOperationException(
                    "The target project name could not be read from SQL Server.");

            #endregion


            #region Verify Captured Project Name

            if (!string.Equals(
                a: databaseProjectName,
                b: expectedProjectName,
                comparisonType: StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    $"The target project name has changed from " +
                    $"'{expectedProjectName}' to '{databaseProjectName}'. " +
                    "Select the CSV again before importing.");
            }

            #endregion
        }

        #endregion


        #region Load Existing Reference Points For Commit

        private static async Task<List<ExistingReferencePoint>>
            LoadExistingReferencePointsForCommitAsync(
                int projectId,
                SqlConnection databaseConnection,
                SqlTransaction transaction)
        {
            #region Define Locked Existing Point Query

            const string existingPointsSql = """
        SELECT
            PN.[PointName_ID],
            PN.[Project_ID],
            PN.[PointName],
            PN.[ReplacementName],
            CR.[Eref],
            CR.[Nref],
            CR.[Href]
        FROM [dbo].[PointName] AS PN WITH (UPDLOCK, HOLDLOCK)
        INNER JOIN [dbo].[CoordinatesReference] AS CR WITH (UPDLOCK, HOLDLOCK)
            ON CR.[PointName_ID] = PN.[PointName_ID]
        WHERE
            PN.[Project_ID] = @Project_ID;
        """;

            #endregion


            #region Load Existing Points

            List<ExistingReferencePoint> existingPoints =
                new();

            await using SqlCommand existingPointsCommand =
                new(
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
                ExistingReferencePoint existingPoint =
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

                        Easting =
                            reader.GetDecimal(
                                i: 4),

                        Northing =
                            reader.GetDecimal(
                                i: 5),

                        Height =
                            reader.GetDecimal(
                                i: 6)
                    };

                existingPoints.Add(
                    item: existingPoint);
            }

            #endregion


            #region Return Existing Points

            return existingPoints;

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
                4;

            eastingParameter.Value =
                importItem.Easting.Value;


            SqlParameter northingParameter =
                insertCoordinateCommand.Parameters.Add(
                    parameterName: "@Nref",
                    sqlDbType: System.Data.SqlDbType.Decimal);

            northingParameter.Precision =
                18;

            northingParameter.Scale =
                4;

            northingParameter.Value =
                importItem.Northing.Value;


            SqlParameter heightParameter =
                insertCoordinateCommand.Parameters.Add(
                    parameterName: "@Href",
                    sqlDbType: System.Data.SqlDbType.Decimal);

            heightParameter.Precision =
                18;

            heightParameter.Scale =
                4;

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


        #region Constructor

        public MainWindow()
        {
            #region Initialise Window Components

            InitializeComponent();

            #endregion


            #region EPPlus License

            gnaT.epplusLicense();

            #endregion


            #region Initialise Project DataGrid

            dgProjects.ItemsSource =
                _projectItems;

            #endregion


            #region Initialise Reference Coordinate Import DataGrid

            dgReferenceCoordinateImport.ItemsSource =
                _referenceImportItems;

            #endregion


            #region Register Window Events

            Loaded +=
                MainWindow_Loaded;

            Closed +=
                MainWindow_Closed;

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





        #endregion


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


        #region Database Connection Test

        private async void btnTestDbConnection_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

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


            #region Test SQL Database Connection

            try
            {
                using SqlConnection sqlConnection =
                    new(
                        connectionString: connectionString);

                await sqlConnection.OpenAsync();

                SaveDatabaseConnectionString(
                    connectionString: connectionString);

                txtDbConnectionStatus.Text =
                    $"Connection successful. " +
                    $"Server: {sqlConnection.DataSource}; " +
                    $"Database: {sqlConnection.Database}";
            }

            #endregion


            #region Handle Connection Errors

            catch (SqlException ex)
            {
                txtDbConnectionStatus.Text =
                    $"SQL connection failed: {ex.Message}";
            }
            catch (InvalidOperationException ex)
            {
                txtDbConnectionStatus.Text =
                    $"Invalid connection configuration: {ex.Message}";
            }
            catch (Exception ex)
            {
                txtDbConnectionStatus.Text =
                    $"Connection test failed: {ex.Message}";
            }

            #endregion


            #region Restore User Interface

            finally
            {
                btnTestDbConnection.IsEnabled =
                    true;
            }

            #endregion
        }

        #endregion


        #region Clear Selected Database Table

        private async void btnClearTables_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Table Selection

            if (cmbClearTables.SelectedItem
                is not ComboBoxItem selectedItem)
            {
                txtDbConnectionStatus.Text =
                    "Select a table.";

                return;
            }

            string selectedTable =
                selectedItem.Content?.ToString()?.Trim()
                ?? string.Empty;

            if (cmbClearTables.SelectedIndex == 0 ||
                string.IsNullOrWhiteSpace(selectedTable))
            {
                txtDbConnectionStatus.Text =
                    "Select a table.";

                return;
            }

            #endregion


            #region Confirm Clear Operation

            string confirmationMessage =
                $"Clear table '{selectedTable}'?\n\nThis cannot be undone.";

            MessageBoxResult confirmation =
                MessageBox.Show(
                    owner: this,
                    messageBoxText: confirmationMessage,
                    caption: "Confirm Clear Table",
                    button: MessageBoxButton.YesNo,
                    icon: MessageBoxImage.Warning,
                    defaultResult: MessageBoxResult.No);

            if (confirmation != MessageBoxResult.Yes)
            {
                txtDbConnectionStatus.Text =
                    "Clear cancelled.";

                return;
            }

            #endregion


            #region Prepare User Interface

            btnClearTables.IsEnabled =
                false;

            cmbClearTables.IsEnabled =
                false;

            txtDbConnectionStatus.Text =
                $"Clearing '{selectedTable}'...";

            #endregion


            try
            {
                #region Clear Selected Table

                await ClearDatabaseTableAsync(
                    tableName: selectedTable);

                #endregion


                #region Reset Application State

                if (selectedTable == "Project")
                {
                    ReleaseActiveProjectLockConnection();

                    _activeProjectId =
                        null;

                    _activeProjectName =
                        string.Empty;

                    txtActiveProject.Text =
                        "No active project";

                    ClearActiveProjectFromRegistry();

                    _projectItems.Clear();

                    ResetReferenceImportState(
                        statusMessage: "No CSV selected.");

                    ResetPrismPairImportState(
                        statusMessage: "No workbook selected.");
                }
                else if (selectedTable == "PointName" ||
                         selectedTable == "CoordinatesReference")
                {
                    ResetReferenceImportState(
                        statusMessage: "No CSV selected.");

                    ResetPrismPairImportState(
                        statusMessage: "No workbook selected.");
                }
                else if (selectedTable == "Track" ||
                         selectedTable == "PrismPairs")
                {
                    ResetPrismPairImportState(
                        statusMessage: "No workbook selected.");
                }

                #endregion


                #region Report Success


                txtDbConnectionStatus.Text =
    $"'{selectedTable}' cleared.";

                cmbClearTables.SelectedIndex =
                    0;

                #endregion
            }
            catch (SqlException ex)
                when (ex.Number == 547)
            {
                #region Report Foreign Key Failure

                txtDbConnectionStatus.Text =
                    $"'{selectedTable}': Related records exist.";

                #endregion
            }
            catch (SqlException ex)
            {
                #region Report SQL Failure

                txtDbConnectionStatus.Text =
                    $"Clear failed: {ex.Message}";

                #endregion
            }
            catch (Exception ex)
            {
                #region Report Clear Failure

                txtDbConnectionStatus.Text =
                    $"Clear failed: {ex.Message}";

                #endregion
            }
            finally
            {
                #region Restore User Interface

                btnClearTables.IsEnabled =
                    true;

                cmbClearTables.IsEnabled =
                    true;

                #endregion
            }
        }


        private async Task ClearDatabaseTableAsync(
            string tableName)
        {


            #region Validate Table Name

            string clearSql =
                tableName switch
                {
                    "CoordinatesReference" =>
                        "DELETE FROM [dbo].[CoordinatesReference];",

                    "PrismPairs" =>
                        """
    SET XACT_ABORT ON;

    BEGIN TRY

        BEGIN TRANSACTION;

        DELETE FROM [dbo].[PrismPairs];
        DELETE FROM [dbo].[Track];

        COMMIT TRANSACTION;

    END TRY

    BEGIN CATCH

        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;

        THROW;

    END CATCH;
    """,

                    _ =>
                        throw new InvalidOperationException(
                            "Invalid table selection.")
                };

            #endregion


            #region Read Database Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string unavailable.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "Database connection string not configured.");
            }

            #endregion


            #region Build Database Connection

            SqlConnectionStringBuilder databaseConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog =
                        TrackGeometryDatabaseName
                };

            #endregion


            #region Execute Clear Operation

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionBuilder.ConnectionString);

            await databaseConnection.OpenAsync();

            await using SqlCommand clearCommand =
                new(
                    cmdText: clearSql,
                    connection: databaseConnection);

            await clearCommand.ExecuteNonQueryAsync();

            #endregion
        }

        #endregion



        #region Database Creation

        private async void btnCreateDb_Click(
            object sender,
            RoutedEventArgs e)
        {
            #region Validate Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                txtDbConnectionStatus.Text =
                    "Enter and test the database connection string first.";

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

                await CreateDatabaseAndTablesAsync(
                    connectionString: connectionString,
                    recreateExistingDatabase: recreateExistingDatabase);

                #endregion


                #region Reset Process Project State After Recreation

                if (recreateExistingDatabase)
                {
                    // The database has been completely recreated.
                    // Any previously held Project_ID is therefore invalid
                    // for this running process.

                    _activeProjectId =
                        null;

                    _activeProjectName =
                        string.Empty;

                    txtActiveProject.Text =
                        "No active project";

                    _projectItems.Clear();

                    // The persisted active project also belongs to the
                    // destroyed database and is no longer valid.

                    ClearActiveProjectFromRegistry();
                }

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

            SqlConnectionStringBuilder masterConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = "master"
                };

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
                    connectionString: masterConnectionBuilder.ConnectionString);

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

            SqlConnectionStringBuilder masterConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = "master"
                };

            #endregion


            #region Create Or Recreate Database

            await using (SqlConnection masterConnection =
                new(
                    connectionString: masterConnectionBuilder.ConnectionString))
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

            SqlConnectionStringBuilder databaseConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = TrackGeometryDatabaseName
                };

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
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_Project_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_Project]
                                PRIMARY KEY CLUSTERED ([Project_ID]),

                            CONSTRAINT [UQ_Project_ProjectName]
                                UNIQUE ([ProjectName])
                        );

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
                       COORDINATES REFERENCE
                       One reference coordinate set per point.
                       Coordinates stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.CoordinatesReference', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[CoordinatesReference]
                        (
                            [PointName_ID] int NOT NULL,
                            [Eref] decimal(18,4) NOT NULL,
                            [Nref] decimal(18,4) NOT NULL,
                            [Href] decimal(18,4) NOT NULL,

                            CONSTRAINT [PK_CoordinatesReference]
                                PRIMARY KEY CLUSTERED ([PointName_ID]),

                            CONSTRAINT [FK_CoordinatesReference_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                    END;


                    /* =============================================================
                       COORDINATES CURRENT
                       One current coordinate set per point.
                       Coordinates stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.CoordinatesCurrent', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[CoordinatesCurrent]
                        (
                            [PointName_ID] int NOT NULL,
                            [UTCtime] datetime2(0) NOT NULL,
                            [E] decimal(18,4) NOT NULL,
                            [N] decimal(18,4) NOT NULL,
                            [H] decimal(18,4) NOT NULL,

                            CONSTRAINT [PK_CoordinatesCurrent]
                                PRIMARY KEY CLUSTERED ([PointName_ID]),

                            CONSTRAINT [FK_CoordinatesCurrent_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_CoordinatesCurrent_UTCtime]
                            ON [dbo].[CoordinatesCurrent] ([UTCtime]);

                    END;


                    /* =============================================================
                       PRISM PAIRS
                       Left and Right points must be different.
                       Same-project validation is performed by the application.

                       Soft deletion:
                           IsDeleted = 0 -> Active
                           IsDeleted = 1 -> Deleted / retired
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.PrismPairs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[PrismPairs]
                        (
                            [PrismPair_ID] int IDENTITY(1,1) NOT NULL,
                            [Left_ID] int NOT NULL,
                            [Right_ID] int NOT NULL,
                            [IsDeleted] bit NOT NULL
                                CONSTRAINT [DF_PrismPairs_IsDeleted]
                                DEFAULT (0),

                            CONSTRAINT [PK_PrismPairs]
                                PRIMARY KEY CLUSTERED ([PrismPair_ID]),

                            CONSTRAINT [UQ_PrismPairs_Left_Right]
                                UNIQUE ([Left_ID], [Right_ID]),

                            CONSTRAINT [CK_PrismPairs_DifferentPoints]
                                CHECK ([Left_ID] <> [Right_ID]),

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
                            [dH] decimal(18,4) NOT NULL,

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
                            [Slew] decimal(18,4) NOT NULL,

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
                            [Top] decimal(18,4) NOT NULL,

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
                       CANT EPOCHS
                       Individual epoch values stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.CantEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[CantEpochs]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [Cant] decimal(18,4) NOT NULL,

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
                       SIGNED twist displacement stored in millimetres over 3 m.
                       Ratio is derived by the reporting software.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ShortTwistEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ShortTwistEpochs]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [ShortTwist] decimal(18,4) NULL,

                            CONSTRAINT [PK_ShortTwistEpochs]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCtime]),

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
                       SIGNED twist displacement stored in millimetres over 15 m.
                       Ratio is derived by the reporting software.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.LongTwistEpochs', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[LongTwistEpochs]
                        (
                            [UTCtime] datetime2(0) NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [LongTwist] decimal(18,4) NULL,

                            CONSTRAINT [PK_LongTwistEpochs]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCtime]),

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
                       One daily mean coordinate set per point.
                       Coordinates stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.CoordinatesDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[CoordinatesDaily]
                        (
                            [UTCdate] date NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [E] decimal(18,4) NOT NULL,
                            [N] decimal(18,4) NOT NULL,
                            [H] decimal(18,4) NOT NULL,

                            CONSTRAINT [PK_CoordinatesDaily]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCdate]),

                            CONSTRAINT [FK_CoordinatesDaily_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_CoordinatesDaily_UTCdate]
                            ON [dbo].[CoordinatesDaily] ([UTCdate]);

                    END;


                    /* =============================================================
                       dH DAILY
                       Daily mean derived from DhEpochs.
                       Value stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.DhDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[DhDaily]
                        (
                            [UTCdate] date NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [dH] decimal(18,4) NOT NULL,

                            CONSTRAINT [PK_DhDaily]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCdate]),

                            CONSTRAINT [FK_DhDaily_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_DhDaily_UTCdate]
                            ON [dbo].[DhDaily] ([UTCdate]);

                    END;


                    /* =============================================================
                       SLEW DAILY
                       Daily mean derived from SlewEpochs.
                       Value stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.SlewDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[SlewDaily]
                        (
                            [UTCdate] date NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [Slew] decimal(18,4) NOT NULL,

                            CONSTRAINT [PK_SlewDaily]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCdate]),

                            CONSTRAINT [FK_SlewDaily_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_SlewDaily_UTCdate]
                            ON [dbo].[SlewDaily] ([UTCdate]);

                    END;


                    /* =============================================================
                       TOP DAILY
                       Daily mean derived from TopEpochs.
                       Value stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.TopDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[TopDaily]
                        (
                            [UTCdate] date NOT NULL,
                            [PointName_ID] int NOT NULL,
                            [Top] decimal(18,4) NOT NULL,

                            CONSTRAINT [PK_TopDaily]
                                PRIMARY KEY CLUSTERED
                                ([PointName_ID], [UTCdate]),

                            CONSTRAINT [FK_TopDaily_PointName]
                                FOREIGN KEY ([PointName_ID])
                                REFERENCES [dbo].[PointName] ([PointName_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_TopDaily_UTCdate]
                            ON [dbo].[TopDaily] ([UTCdate]);

                    END;


                    /* =============================================================
                       CANT DAILY
                       Daily mean derived from CantEpochs.
                       Value stored in metres.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.CantDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[CantDaily]
                        (
                            [UTCdate] date NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [Cant] decimal(18,4) NOT NULL,

                            CONSTRAINT [PK_CantDaily]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCdate]),

                            CONSTRAINT [FK_CantDaily_PrismPairs]
                                FOREIGN KEY ([PrismPair_ID])
                                REFERENCES [dbo].[PrismPairs] ([PrismPair_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_CantDaily_UTCdate]
                            ON [dbo].[CantDaily] ([UTCdate]);

                    END;


                    /* =============================================================
                       SHORT TWIST DAILY
                       Daily mean of SIGNED ShortTwistEpochs displacement.
                       Value stored in millimetres over a 3 m base.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.ShortTwistDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[ShortTwistDaily]
                        (
                            [UTCdate] date NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [ShortTwist] decimal(18,4) NULL,

                            CONSTRAINT [PK_ShortTwistDaily]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCdate]),

                            CONSTRAINT [FK_ShortTwistDaily_PrismPairs]
                                FOREIGN KEY ([PrismPair_ID])
                                REFERENCES [dbo].[PrismPairs] ([PrismPair_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_ShortTwistDaily_UTCdate]
                            ON [dbo].[ShortTwistDaily] ([UTCdate]);

                    END;


                    /* =============================================================
                       LONG TWIST DAILY
                       Daily mean of SIGNED LongTwistEpochs displacement.
                       Value stored in millimetres over a 15 m base.
                       ============================================================= */

                    IF OBJECT_ID(N'dbo.LongTwistDaily', N'U') IS NULL
                    BEGIN

                        CREATE TABLE [dbo].[LongTwistDaily]
                        (
                            [UTCdate] date NOT NULL,
                            [PrismPair_ID] int NOT NULL,
                            [LongTwist] decimal(18,4) NULL,

                            CONSTRAINT [PK_LongTwistDaily]
                                PRIMARY KEY CLUSTERED
                                ([PrismPair_ID], [UTCdate]),

                            CONSTRAINT [FK_LongTwistDaily_PrismPairs]
                                FOREIGN KEY ([PrismPair_ID])
                                REFERENCES [dbo].[PrismPairs] ([PrismPair_ID])
                                ON DELETE NO ACTION
                                ON UPDATE NO ACTION
                        );

                        CREATE INDEX [IX_LongTwistDaily_UTCdate]
                            ON [dbo].[LongTwistDaily] ([UTCdate]);

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
                        databaseConnectionBuilder.ConnectionString))
            {
                await databaseConnection.OpenAsync();

                await using SqlCommand createTablesCommand =
                    new(
                        cmdText: createTablesSql,
                        connection: databaseConnection);

                await createTablesCommand.ExecuteNonQueryAsync();
            }


            await EnsureTrackAndPrismPairSchemaAsync(
                connectionString:
                    databaseConnectionBuilder.ConnectionString);

            #endregion



        }


        #region Track And Prism Pair Schema Upgrade

        private static async Task EnsureTrackAndPrismPairSchemaAsync(
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


            #region Define Schema Upgrade Phase One

            // IMPORTANT:
            //
            // Track_ID and PairOrder must be ADDED in a separate SQL batch from
            // any SQL statements that subsequently reference those columns.
            //
            // SQL Server compiles an entire batch before executing it. Therefore
            // adding a column and then referencing that new column later in the
            // same batch can produce:
            //
            //     Invalid column name 'PairOrder'
            //
            // even though the ALTER TABLE ADD statement appears earlier.

            const string phaseOneSql = """
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
           Validate existing table before migration.
           ============================================================= */

        IF OBJECT_ID(N'dbo.PrismPairs', N'U') IS NULL
        BEGIN

            THROW 50001,
                'PrismPairs table does not exist.',
                1;

        END;


        /* -------------------------------------------------------------
           Reject migration of populated legacy PrismPairs.

           Track_ID and PairOrder cannot safely be inferred from the
           old PrismPairs records.
           ------------------------------------------------------------- */

        IF
        (
            COL_LENGTH(
                'dbo.PrismPairs',
                'Track_ID') IS NULL

            OR

            COL_LENGTH(
                'dbo.PrismPairs',
                'PairOrder') IS NULL
        )
        AND EXISTS
        (
            SELECT 1
            FROM [dbo].[PrismPairs]
        )
        BEGIN

            THROW 50002,
                'Legacy PrismPairs contains existing data. Automatic Track migration is blocked.',
                1;

        END;


        /* -------------------------------------------------------------
           Add Track_ID.
           ------------------------------------------------------------- */

        IF COL_LENGTH(
            'dbo.PrismPairs',
            'Track_ID') IS NULL
        BEGIN

            ALTER TABLE [dbo].[PrismPairs]
            ADD [Track_ID] int NULL;

        END;


        /* -------------------------------------------------------------
           Add PairOrder.
           ------------------------------------------------------------- */

        IF COL_LENGTH(
            'dbo.PrismPairs',
            'PairOrder') IS NULL
        BEGIN

            ALTER TABLE [dbo].[PrismPairs]
            ADD [PairOrder] int NULL;

        END;
        """;

            #endregion


            #region Define Schema Upgrade Phase Two

            // This is deliberately a SECOND SQL batch.
            //
            // When SQL Server compiles this batch, Track_ID and PairOrder already
            // exist because Phase One has completed on the same SQL connection
            // and within the same transaction.

            const string phaseTwoSql = """
        /* =============================================================
           VERIFY NEW PRISM PAIR COLUMNS
           ============================================================= */

        IF COL_LENGTH(
            'dbo.PrismPairs',
            'Track_ID') IS NULL
        BEGIN

            THROW 50003,
                'Track_ID was not created in PrismPairs.',
                1;

        END;


        IF COL_LENGTH(
            'dbo.PrismPairs',
            'PairOrder') IS NULL
        BEGIN

            THROW 50004,
                'PairOrder was not created in PrismPairs.',
                1;

        END;


        /* =============================================================
           VERIFY MIGRATION DATA STATE
           ============================================================= */

        IF EXISTS
        (
            SELECT 1
            FROM [dbo].[PrismPairs]
            WHERE
                [Track_ID] IS NULL
                OR
                [PairOrder] IS NULL
        )
        BEGIN

            THROW 50005,
                'PrismPairs contains records without Track_ID or PairOrder.',
                1;

        END;


        /* =============================================================
           MAKE NEW COLUMNS MANDATORY
           ============================================================= */

        IF EXISTS
        (
            SELECT 1
            FROM sys.columns
            WHERE
                [object_id] =
                    OBJECT_ID(N'dbo.PrismPairs')
                AND [name] =
                    N'Track_ID'
                AND [is_nullable] =
                    1
        )
        BEGIN

            ALTER TABLE [dbo].[PrismPairs]
            ALTER COLUMN [Track_ID] int NOT NULL;

        END;


        IF EXISTS
        (
            SELECT 1
            FROM sys.columns
            WHERE
                [object_id] =
                    OBJECT_ID(N'dbo.PrismPairs')
                AND [name] =
                    N'PairOrder'
                AND [is_nullable] =
                    1
        )
        BEGIN

            ALTER TABLE [dbo].[PrismPairs]
            ALTER COLUMN [PairOrder] int NOT NULL;

        END;


        /* =============================================================
           TRACK FOREIGN KEY
           ============================================================= */

        IF NOT EXISTS
        (
            SELECT 1
            FROM sys.foreign_keys
            WHERE
                [name] =
                    N'FK_PrismPairs_Track'
                AND [parent_object_id] =
                    OBJECT_ID(N'dbo.PrismPairs')
        )
        BEGIN

            ALTER TABLE [dbo].[PrismPairs]
            ADD CONSTRAINT [FK_PrismPairs_Track]
                FOREIGN KEY ([Track_ID])
                REFERENCES [dbo].[Track] ([Track_ID])
                ON DELETE NO ACTION
                ON UPDATE NO ACTION;

        END;


        /* =============================================================
           PAIR ORDER CHECK
           ============================================================= */

        IF NOT EXISTS
        (
            SELECT 1
            FROM sys.check_constraints
            WHERE
                [name] =
                    N'CK_PrismPairs_PairOrder'
                AND [parent_object_id] =
                    OBJECT_ID(N'dbo.PrismPairs')
        )
        BEGIN

            ALTER TABLE [dbo].[PrismPairs]
            ADD CONSTRAINT [CK_PrismPairs_PairOrder]
                CHECK ([PairOrder] > 0);

        END;


        /* =============================================================
           PAIR ORDER UNIQUE WITHIN TRACK
           ============================================================= */

        IF NOT EXISTS
        (
            SELECT 1
            FROM sys.key_constraints
            WHERE
                [name] =
                    N'UQ_PrismPairs_Track_PairOrder'
                AND [parent_object_id] =
                    OBJECT_ID(N'dbo.PrismPairs')
        )
        BEGIN

            ALTER TABLE [dbo].[PrismPairs]
            ADD CONSTRAINT [UQ_PrismPairs_Track_PairOrder]
                UNIQUE ([Track_ID], [PairOrder]);

        END;


        /* =============================================================
           TRACK LOOKUP INDEX
           ============================================================= */

        IF NOT EXISTS
        (
            SELECT 1
            FROM sys.indexes
            WHERE
                [object_id] =
                    OBJECT_ID(N'dbo.PrismPairs')
                AND [name] =
                    N'IX_PrismPairs_Track_ID'
        )
        BEGIN

            CREATE INDEX [IX_PrismPairs_Track_ID]
                ON [dbo].[PrismPairs] ([Track_ID]);

        END;
        """;

            #endregion


            #region Open Schema Upgrade Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString: connectionString);

            await databaseConnection.OpenAsync();

            #endregion


            #region Begin Schema Upgrade Transaction

            using SqlTransaction transaction =
                databaseConnection.BeginTransaction(
                    iso:
                        System.Data.IsolationLevel.Serializable);

            bool transactionCommitted =
                false;

            #endregion


            try
            {
                #region Apply Schema Upgrade Phase One

                await using (SqlCommand phaseOneCommand =
                    new(
                        cmdText: phaseOneSql,
                        connection: databaseConnection,
                        transaction: transaction))
                {
                    await phaseOneCommand.ExecuteNonQueryAsync();
                }

                #endregion


                #region Apply Schema Upgrade Phase Two

                await using (SqlCommand phaseTwoCommand =
                    new(
                        cmdText: phaseTwoSql,
                        connection: databaseConnection,
                        transaction: transaction))
                {
                    await phaseTwoCommand.ExecuteNonQueryAsync();
                }

                #endregion


                #region Commit Schema Upgrade

                transaction.Commit();

                transactionCommitted =
                    true;

                #endregion
            }
            catch
            {
                #region Roll Back Schema Upgrade

                if (!transactionCommitted)
                {
                    try
                    {
                        transaction.Rollback();
                    }
                    catch
                    {
                        // Preserve the original schema-upgrade exception.
                    }
                }

                throw;

                #endregion
            }
        }

        #endregion




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


            #region Read Database Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "Database connection string has not been configured.");
            }

            #endregion


            #region Build Dedicated Lock Connection String

            SqlConnectionStringBuilder lockConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = TrackGeometryDatabaseName,

                    // The physical SQL session must terminate when this connection
                    // is disposed so its Session-owned application lock is released.
                    Pooling = false
                };

            #endregion


            #region Open Dedicated Lock Connection

            SqlConnection lockConnection =
                new(
                    connectionString:
                        lockConnectionBuilder.ConnectionString);

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


            #region Read Database Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "Database connection string has not been configured.");
            }

            #endregion


            #region Build Dedicated Lock Connection String

            SqlConnectionStringBuilder lockConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = TrackGeometryDatabaseName,

                    // The Exclusive application lock belongs to this SQL session.
                    // Pooling is disabled so disposal terminates the physical
                    // session and releases the lock deterministically.
                    Pooling = false
                };

            #endregion


            #region Open Dedicated Lock Connection

            SqlConnection lockConnection =
                new(
                    connectionString:
                        lockConnectionBuilder.ConnectionString);

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

                _activeProjectLockProjectId =
                    startupProjectId;

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

            _activeProjectLockProjectId =
                null;

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




        #region Add Project

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

            btnAddProject.IsEnabled =
                false;

            txtProjectManagementStatus.Text =
                $"Checking project '{projectName}'...";

            #endregion


            try
            {
                #region Add Or Restore Project

                bool projectAddedOrRestored =
                    await AddOrRestoreProjectAsync(
                        projectName: projectName);

                #endregion


                #region Refresh Project List

                if (projectAddedOrRestored)
                {
                    txtNewProjectName.Clear();

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
                btnAddProject.IsEnabled =
                    true;
            }

            #endregion
        }


        private async Task<bool> AddOrRestoreProjectAsync(
            string projectName)
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

            #endregion


            #region Read Database Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "Database connection string has not been configured.");
            }

            #endregion


            #region Build Track Geometry Database Connection String

            SqlConnectionStringBuilder databaseConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = TrackGeometryDatabaseName
                };

            #endregion


            #region Open Database Connection

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionBuilder.ConnectionString);

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
                txtProjectManagementStatus.Text =
                    $"Project '{validatedProjectName}' already exists.";

                return false;
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
            SET [IsDeleted] = 0
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
            [IsDeleted]
        )
        VALUES
        (
            @ProjectName,
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

                #endregion


                #region Transfer Active Project Lock Ownership

                // The newly opened SQL connection now becomes the dedicated Shared
                // lock connection for this application instance.

                _activeProjectLockConnection =
                    newProjectLockConnection;

                _activeProjectLockProjectId =
                    selectedProject.Project_ID;

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

                ResetPrismPairImportState(
                    statusMessage:
                        $"Active project changed to '{_activeProjectName}'. " +
                        "Select a Track Geometry workbook for this project.");

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


            #region Read Database Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "Database connection string has not been configured.");
            }

            #endregion


            #region Build Database Connection String

            SqlConnectionStringBuilder databaseConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = TrackGeometryDatabaseName
                };

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
                        databaseConnectionBuilder.ConnectionString);

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

            #region Read Database Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "Database connection string has not been configured.");
            }

            #endregion


            #region Build Track Geometry Database Connection String

            SqlConnectionStringBuilder databaseConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = TrackGeometryDatabaseName
                };

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
                        databaseConnectionBuilder.ConnectionString);

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
            #region Read Database Connection String

            string connectionString =
                txtDbConnectionString.Text?.Trim()
                ?? throw new InvalidOperationException(
                    "Database connection string control returned null.");

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException(
                    "Database connection string has not been configured.");
            }

            #endregion


            #region Build Track Geometry Database Connection String

            SqlConnectionStringBuilder databaseConnectionBuilder =
                new(
                    connectionString: connectionString)
                {
                    InitialCatalog = TrackGeometryDatabaseName
                };

            #endregion


            #region Define Project Query

            const string loadProjectsSql = """
                SELECT
                    [Project_ID],
                    [ProjectName],
                    [IsDeleted]
                FROM [dbo].[Project]
                ORDER BY
                    [ProjectName];
                """;

            #endregion


            #region Clear Existing Project List

            _projectItems.Clear();

            #endregion


            #region Load Projects From Database

            await using SqlConnection databaseConnection =
                new(
                    connectionString:
                        databaseConnectionBuilder.ConnectionString);

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

                bool isDeleted =
                    reader.GetBoolean(
                        i: 2);

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

            #endregion
        }

        #endregion
    }
}