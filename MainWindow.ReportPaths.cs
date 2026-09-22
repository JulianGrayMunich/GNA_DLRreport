#region System Preparation

using System;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;

#endregion

namespace GNA_DLRreport
{
    public partial class MainWindow
    {
        #region Report Selection State

        private const int ReportPathMaximumLength = 1000;
        private const int ReportNameMaximumLength = 200;
        private const string ReportTemplateExtension = ".docx";
        private readonly SemaphoreSlim _reportSelectionGate = new(initialCount: 1, maxCount: 1);
        private int _reportSelectionVersion;
        private int _reportOutputChangeVersion;
        private int? _reportSelectionProjectId;
        private string _reportSelectionConnectionString = string.Empty;
        private string _savedReportName = string.Empty;

        #endregion

        #region Report Selection Context

        private void InvalidateReportSelections()
        {
            if (txtReportTemplatePath is null ||
                (_reportSelectionProjectId == _activeProjectId && _activeProjectId.HasValue))
            {
                return;
            }

            _reportSelectionVersion++;
            _reportSelectionProjectId = null;
            _reportSelectionConnectionString = string.Empty;
            txtReportTemplatePath.Clear();
            txtReportOutputFolder.Clear();
            txtReportName.Clear();
            _savedReportName = string.Empty;
            SetReportSelectionButtons(enabled: false);
            UpdateReportGenerationAvailability();
        }

        private void SetReportSelectionButtons(bool enabled)
        {
            btnSelectReportTemplate.IsEnabled = enabled;
            btnSelectReportOutputFolder.IsEnabled = enabled;
            txtReportName.IsEnabled = enabled;
        }

        private bool IsCurrentReportSelectionContext(int projectId, string connectionString, int version)
        {
            return _activeProjectId == projectId &&
                _reportSelectionVersion == version &&
                string.Equals(a: GetTrackGeometryConnectionString(), b: connectionString,
                    comparisonType: StringComparison.Ordinal);
        }

        private async void tabReportGeneration_Selected(object sender, RoutedEventArgs e)
        {
            if (ReferenceEquals(objA: e.OriginalSource, objB: tabReportGeneration) && IsLoaded)
            {
                await RefreshReportSelectionsAsync();
            }
        }

        #endregion

        #region Report Selection Schema

        private static async Task EnsureReportSelectionSchemaAsync(SqlConnection connection)
        {
            // Serialize the missing-column checks across application instances.
            // Keep this command separate from commands which reference new columns.
            const string sql = """
                SET XACT_ABORT ON;
                BEGIN TRY
                    BEGIN TRANSACTION;
                    DECLARE @LockResult int;
                    EXEC @LockResult = sys.sp_getapplock
                        @Resource = N'GNA_DLRreport:ReportSelectionSchema',
                        @LockMode = N'Exclusive',
                        @LockOwner = N'Transaction',
                        @LockTimeout = 10000;
                    IF @LockResult < 0
                        THROW 51039, 'Unable to lock report selection schema.', 1;
                    IF OBJECT_ID(N'dbo.Project', N'U') IS NULL
                        THROW 51039, 'The Project table is unavailable.', 1;
                    IF COL_LENGTH(N'dbo.Project', N'ReportTemplatePath') IS NULL
                        ALTER TABLE [dbo].[Project] ADD [ReportTemplatePath] nvarchar(1000) NULL;
                    IF COL_LENGTH(N'dbo.Project', N'DefaultReportOutputPath') IS NULL
                        ALTER TABLE [dbo].[Project] ADD [DefaultReportOutputPath] nvarchar(1000) NULL;
                    IF COL_LENGTH(N'dbo.Project', N'ReportName') IS NULL
                        ALTER TABLE [dbo].[Project] ADD [ReportName] nvarchar(200) NULL;
                    IF EXISTS
                    (
                        SELECT 1 FROM sys.columns
                        WHERE object_id = OBJECT_ID(N'dbo.Project')
                          AND name IN (N'ReportTemplatePath', N'DefaultReportOutputPath')
                          AND (system_type_id <> TYPE_ID(N'nvarchar')
                               OR (max_length <> -1 AND max_length < 2000)
                               OR is_computed = 1)
                    )
                        THROW 51039, 'Report path columns must support nvarchar(1000).', 1;
                    IF EXISTS
                    (
                        SELECT 1 FROM sys.columns
                        WHERE object_id = OBJECT_ID(N'dbo.Project')
                          AND name = N'ReportName'
                          AND (system_type_id <> TYPE_ID(N'nvarchar')
                               OR (max_length <> -1 AND max_length < 400)
                               OR is_computed = 1)
                    )
                        THROW 51039, 'ReportName must support nvarchar(200).', 1;
                    COMMIT TRANSACTION;
                END TRY
                BEGIN CATCH
                    IF XACT_STATE() <> 0 ROLLBACK TRANSACTION;
                    THROW;
                END CATCH;
                """;

            await using SqlCommand command = new(cmdText: sql, connection: connection);
            await command.ExecuteNonQueryAsync();
        }

        #endregion

        #region Load Persisted Report Selections

        private async Task RefreshReportSelectionsAsync()
        {
            int version = ++_reportSelectionVersion;
            _reportSelectionProjectId = null;
            txtReportTemplatePath.Clear();
            txtReportOutputFolder.Clear();
            txtReportName.Clear();
            _savedReportName = string.Empty;
            SetReportSelectionButtons(enabled: false);
            UpdateReportGenerationAvailability();

            if (!_activeProjectId.HasValue)
            {
                return;
            }

            int projectId = _activeProjectId.Value;
            bool gateAcquired = false;
            try
            {
                string connectionString = GetTrackGeometryConnectionString();
                await _reportSelectionGate.WaitAsync();
                gateAcquired = true;
                if (!IsCurrentReportSelectionContext(projectId: projectId,
                    connectionString: connectionString, version: version))
                {
                    return;
                }

                await using SqlConnection connection = new(connectionString: connectionString);
                await connection.OpenAsync();
                await EnsureReportSelectionSchemaAsync(connection: connection);

                const string sql = """
                    SELECT [ReportTemplatePath], [DefaultReportOutputPath], [ReportName]
                    FROM [dbo].[Project]
                    WHERE [Project_ID] = @Project_ID AND [IsDeleted] = 0;
                    """;
                await using SqlCommand command = new(cmdText: sql, connection: connection);
                command.Parameters.Add(parameterName: "@Project_ID", sqlDbType: SqlDbType.Int).Value = projectId;
                await using SqlDataReader reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync())
                {
                    throw new InvalidOperationException(message: "The active project is unavailable.");
                }

                string templatePath = reader.IsDBNull(i: 0) ? string.Empty : reader.GetString(i: 0);
                string outputFolder = reader.IsDBNull(i: 1) ? string.Empty : reader.GetString(i: 1);
                string reportName = reader.IsDBNull(i: 2) ? string.Empty : reader.GetString(i: 2);
                if (!IsCurrentReportSelectionContext(projectId: projectId,
                    connectionString: connectionString, version: version))
                {
                    return;
                }

                _reportSelectionProjectId = projectId;
                _reportSelectionConnectionString = connectionString;
                txtReportTemplatePath.Text = templatePath;
                txtReportOutputFolder.Text = outputFolder;
                _savedReportName = reportName;
                txtReportName.Text = reportName;
                UpdateReportGenerationAvailability();
                SetReportSelectionButtons(enabled: true);
                txtReportGenerationStatus.Text = "Report selections loaded. Browse to select and save each path.";
            }
            catch (Exception ex)
            {
                if (version == _reportSelectionVersion)
                {
                    txtReportGenerationStatus.Text = $"Unable to load report selections: {ex.Message}";
                }
            }
            finally
            {
                if (gateAcquired)
                {
                    _reportSelectionGate.Release();
                }
            }
        }

        #endregion

        #region Persist Report Name

        private void txtReportName_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            UpdateReportGenerationAvailability();
        }

        private async void txtReportName_LostFocus(object sender, RoutedEventArgs e)
        {
            if (_reportSelectionProjectId != _activeProjectId || !_activeProjectId.HasValue)
            {
                return;
            }

            try
            {
                await SaveReportNameAsync();
            }
            catch (Exception ex)
            {
                txtReportGenerationStatus.Text = $"Report name was not saved: {ex.Message}";
            }
        }

        private async Task SaveReportNameAsync()
        {
            string reportName = txtReportName.Text.Trim();
            if (reportName.Length > ReportNameMaximumLength)
            {
                throw new InvalidOperationException(
                    "Report Name must contain no more than 200 characters.");
            }
            if (string.Equals(a: reportName, b: _savedReportName,
                comparisonType: StringComparison.Ordinal))
            {
                return;
            }

            int projectId = _reportSelectionProjectId ??
                throw new InvalidOperationException("Select an active project.");
            int version = _reportSelectionVersion;
            string connectionString = _reportSelectionConnectionString;
            await _reportSelectionGate.WaitAsync();
            try
            {
                if (!IsCurrentReportSelectionContext(
                    projectId: projectId,
                    connectionString: connectionString,
                    version: version))
                {
                    throw new InvalidOperationException(
                        "The active project changed before the report name could be saved.");
                }

                const string sql = """
                    UPDATE [dbo].[Project]
                    SET [ReportName] = NULLIF(@ReportName, N'')
                    WHERE [Project_ID] = @Project_ID AND [IsDeleted] = 0;
                    """;
                await using SqlConnection connection = new(
                    connectionString: connectionString);
                await connection.OpenAsync();
                await EnsureReportSelectionSchemaAsync(connection: connection);
                await using SqlCommand command = new(cmdText: sql, connection: connection);
                command.Parameters.Add(parameterName: "@Project_ID", sqlDbType: SqlDbType.Int)
                    .Value = projectId;
                command.Parameters.Add(parameterName: "@ReportName",
                    sqlDbType: SqlDbType.NVarChar,
                    size: ReportNameMaximumLength).Value = reportName;
                if (await command.ExecuteNonQueryAsync() != 1)
                {
                    throw new InvalidOperationException(
                        "The active project is unavailable for saving.");
                }

                _savedReportName = reportName;
                if (string.Equals(a: txtReportName.Text.Trim(), b: reportName,
                    comparisonType: StringComparison.Ordinal))
                {
                    txtReportGenerationStatus.Text = "Report name saved for the active project.";
                }
            }
            finally
            {
                _reportSelectionGate.Release();
                UpdateReportGenerationAvailability();
            }
        }

        #endregion

        #region Report Selection Dialogs

        private async void btnSelectReportTemplate_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new()
            {
                Title = "Select Word Report Template",
                Filter = "Word documents (*.docx)|*.docx",
                DefaultExt = ReportTemplateExtension,
                CheckFileExists = true,
                CheckPathExists = true,
                Multiselect = false
            };
            if (File.Exists(path: txtReportTemplatePath.Text))
            {
                dialog.FileName = txtReportTemplatePath.Text;
            }
            if (dialog.ShowDialog(owner: this) == true)
            {
                await SaveReportSelectionAsync(path: dialog.FileName, isTemplate: true);
            }
        }

        private async void btnSelectReportOutputFolder_Click(object sender, RoutedEventArgs e)
        {
            OpenFolderDialog dialog = new()
            {
                Title = "Select Report Output Folder",
                Multiselect = false
            };
            if (Directory.Exists(path: txtReportOutputFolder.Text))
            {
                dialog.InitialDirectory = txtReportOutputFolder.Text;
            }
            if (dialog.ShowDialog(owner: this) == true)
            {
                await SaveReportSelectionAsync(path: dialog.FolderName, isTemplate: false);
            }
        }

        #endregion

        #region Persist Selected Report Path

        private async Task SaveReportSelectionAsync(string path, bool isTemplate)
        {
            int version = _reportSelectionVersion;
            bool gateAcquired = false;
            try
            {
                int projectId = _reportSelectionProjectId
                    ?? throw new InvalidOperationException(message: "Load an active project before selecting report paths.");
                string connectionString = _reportSelectionConnectionString;
                string validatedPath = ValidateReportSelectionPath(path: path, isTemplate: isTemplate);
                SetReportSelectionButtons(enabled: false);
                await _reportSelectionGate.WaitAsync();
                gateAcquired = true;
                if (!IsCurrentReportSelectionContext(projectId: projectId,
                    connectionString: connectionString, version: version))
                {
                    return;
                }

                // Write only the selected field, preserving the other persisted setting.
                const string templateSql = """
                    UPDATE [dbo].[Project] SET [ReportTemplatePath] = @Path
                    WHERE [Project_ID] = @Project_ID AND [IsDeleted] = 0;
                    """;
                const string outputSql = """
                    UPDATE [dbo].[Project] SET [DefaultReportOutputPath] = @Path
                    WHERE [Project_ID] = @Project_ID AND [IsDeleted] = 0;
                    """;
                await using SqlConnection connection = new(connectionString: connectionString);
                await connection.OpenAsync();
                await using SqlCommand command = new(cmdText: isTemplate ? templateSql : outputSql, connection: connection);
                command.Parameters.Add(parameterName: "@Project_ID", sqlDbType: SqlDbType.Int).Value = projectId;
                command.Parameters.Add(parameterName: "@Path", sqlDbType: SqlDbType.NVarChar,
                    size: ReportPathMaximumLength).Value = validatedPath;
                int affectedRows = await command.ExecuteNonQueryAsync();
                if (affectedRows != 1)
                {
                    throw new InvalidOperationException(message: "The active project is no longer available for saving.");
                }
                if (!IsCurrentReportSelectionContext(projectId: projectId,
                    connectionString: connectionString, version: version))
                {
                    return;
                }

                if (isTemplate)
                {
                    txtReportTemplatePath.Text = validatedPath;
                }
                else
                {
                    _reportOutputChangeVersion++;
                    txtReportOutputFolder.Text = validatedPath;
                    if (dgProjects.SelectedItem is ProjectConfigurationItem selectedProject &&
                        selectedProject.Project_ID == projectId)
                    {
                        txtProjectOutputPath.Text = validatedPath;
                    }
                }
                UpdateReportGenerationAvailability();
                txtReportGenerationStatus.Text = isTemplate
                    ? "Word report template saved for the active project."
                    : "Report output folder saved for the active project.";
            }
            catch (Exception ex)
            {
                if (version == _reportSelectionVersion)
                {
                    txtReportGenerationStatus.Text = $"Report selection was not saved: {ex.Message}";
                }
            }
            finally
            {
                if (gateAcquired)
                {
                    _reportSelectionGate.Release();
                }
                if (version == _reportSelectionVersion)
                {
                    SetReportSelectionButtons(enabled: _reportSelectionProjectId == _activeProjectId &&
                        _activeProjectId.HasValue);
                }
            }
        }

        #endregion

        #region Report Path Validation And Filename Convention

        private static string ValidateReportSelectionPath(string path, bool isTemplate)
        {
            string validatedPath = path?.Trim()
                ?? throw new ArgumentNullException(paramName: nameof(path));
            if (validatedPath.Length == 0 || validatedPath.Length > ReportPathMaximumLength ||
                !Path.IsPathFullyQualified(path: validatedPath))
            {
                throw new ArgumentException(message: "Select an absolute path of no more than 1000 characters.",
                    paramName: nameof(path));
            }
            if (isTemplate)
            {
                if (!string.Equals(a: Path.GetExtension(path: validatedPath), b: ReportTemplateExtension,
                    comparisonType: StringComparison.OrdinalIgnoreCase) || !File.Exists(path: validatedPath))
                {
                    throw new ArgumentException(message: "Select an existing .docx Word template.", paramName: nameof(path));
                }
            }
            else if (!Directory.Exists(path: validatedPath))
            {
                throw new ArgumentException(message: "Select an existing report output folder.", paramName: nameof(path));
            }
            return validatedPath;
        }

        private static string CreateReportFileName(string reportName, DateTime reportEndDate)
        {
            string validatedName = reportName.Trim();
            if (validatedName.EndsWith(value: ReportTemplateExtension,
                comparisonType: StringComparison.OrdinalIgnoreCase))
            {
                validatedName = validatedName[..^ReportTemplateExtension.Length].TrimEnd();
            }
            if (validatedName.Length == 0 ||
                validatedName.IndexOfAny(anyOf: Path.GetInvalidFileNameChars()) >= 0 ||
                validatedName.EndsWith(value: ".", comparisonType: StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    "Report Name cannot be used as a Windows filename.");
            }
            return validatedName + "_" +
                reportEndDate.ToString(format: "yyyyMMdd", provider: CultureInfo.InvariantCulture) +
                ReportTemplateExtension;
        }

        private void UpdateReportGenerationAvailability()
        {
            if (btnReportGeneration is null || txtReportTemplatePath is null ||
                txtReportOutputFolder is null || txtReportName is null ||
                dpReportStartDate is null ||
                dpReportEndDate is null)
            {
                return;
            }

            txtReportActiveProject.Text = _activeProjectId.HasValue &&
                !string.IsNullOrWhiteSpace(value: _activeProjectName)
                ? $"Project: {_activeProjectName}"
                : "Project: —";

            btnReportGeneration.IsEnabled = _activeProjectId.HasValue &&
                _reportSelectionProjectId == _activeProjectId &&
                !string.IsNullOrWhiteSpace(value: txtReportName.Text) &&
                dpReportStartDate.SelectedDate.HasValue &&
                dpReportEndDate.SelectedDate.HasValue &&
                dpReportStartDate.SelectedDate.Value.Date < dpReportEndDate.SelectedDate.Value.Date &&
                File.Exists(path: txtReportTemplatePath.Text) &&
                Directory.Exists(path: txtReportOutputFolder.Text);
        }

        #endregion
    }
}
