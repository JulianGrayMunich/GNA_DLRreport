using System;
using System.Windows;
using System.Threading.Tasks;

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

        #endregion

        #region Database Configuration

        private const string TrackGeometryDatabaseName =
            "DBTrackGeometry";

        #endregion


        #region Constructor

        public MainWindow()
        {
            InitializeComponent();

            Loaded += MainWindow_Loaded;
        }

        #endregion


        #region Window Initialisation

        private void MainWindow_Loaded(
            object sender,
            RoutedEventArgs e)
        {
            #region Load Persisted Configuration

            LoadDatabaseConnectionString();

            #endregion
        }

        #endregion


        #region Registry Persistence

        private void LoadDatabaseConnectionString()
        {
            #region Initialise Connection String

            string connectionString = string.Empty;

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

            txtDbConnectionString.Text = connectionString;

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

            btnTestDbConnection.IsEnabled = false;

            txtDbConnectionStatus.Text =
                "Testing database connection...";

            #endregion


            #region Test SQL Database Connection

            try
            {
                using SqlConnection sqlConnection =
                    new(connectionString: connectionString);

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
                btnTestDbConnection.IsEnabled = true;
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

            btnCreateDb.IsEnabled = false;
            btnTestDbConnection.IsEnabled = false;

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

                bool recreateExistingDatabase = false;

                if (databaseExists)
                {
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

                    recreateExistingDatabase = true;
                }

                #endregion


                #region Final Database Recreation Confirmation

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

                recreateExistingDatabase = true;

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
                btnCreateDb.IsEnabled = true;
                btnTestDbConnection.IsEnabled = true;
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
                new(connectionString: connectionString)
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
                new(connectionString: masterConnectionBuilder.ConnectionString);

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
                new(connectionString: connectionString)
                {
                    InitialCatalog = "master"
                };

            #endregion


            #region Create Or Recreate Database

            await using (SqlConnection masterConnection =
                new(connectionString: masterConnectionBuilder.ConnectionString))
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
                new(connectionString: connectionString)
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
           ============================================================= */

        IF OBJECT_ID(N'dbo.Project', N'U') IS NULL
        BEGIN

            CREATE TABLE [dbo].[Project]
            (
                [Project_ID] int IDENTITY(1,1) NOT NULL,
                [ProjectName] nvarchar(200) NOT NULL,

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
           ============================================================= */

        IF OBJECT_ID(N'dbo.PointName', N'U') IS NULL
        BEGIN

            CREATE TABLE [dbo].[PointName]
            (
                [PointName_ID] int IDENTITY(1,1) NOT NULL,
                [PointName] nvarchar(50) NOT NULL,
                [ReplacementName] nvarchar(50) NOT NULL,
                [Project_ID] int NOT NULL,

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
           ============================================================= */

        IF OBJECT_ID(N'dbo.PrismPairs', N'U') IS NULL
        BEGIN

            CREATE TABLE [dbo].[PrismPairs]
            (
                [PrismPair_ID] int IDENTITY(1,1) NOT NULL,
                [Left_ID] int NOT NULL,
                [Right_ID] int NOT NULL,

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
                new(connectionString: databaseConnectionBuilder.ConnectionString))
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




    }
}