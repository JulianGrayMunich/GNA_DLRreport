using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Windows;

#region WPF Control Namespaces

using System.Windows.Controls;

#endregion


using Microsoft.Data.SqlClient;
using Microsoft.Win32;


namespace GNA_DLRreport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
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


        #region Constructor

        public MainWindow()
        {
            #region Initialise Window Components

            InitializeComponent();

            #endregion


            #region Initialise Project DataGrid

            dgProjects.ItemsSource =
                _projectItems;

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
                    connectionString: databaseConnectionBuilder.ConnectionString))
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

                txtProjectManagementStatus.Text =
                    $"Active project set to '{databaseProjectName}'.";

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