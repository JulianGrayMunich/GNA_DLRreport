#region System Preparation

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Xml.Linq;
using Microsoft.Data.SqlClient;

#endregion

namespace GNA_DLRreport
{
    public partial class MainWindow
    {
        #region Report Generation Open XML Contract

        private static readonly XNamespace WordNs =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace RelationshipNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageRelationshipNs =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace DrawingNs =
            "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace WordDrawingNs =
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        private static readonly XNamespace PictureNs =
            "http://schemas.openxmlformats.org/drawingml/2006/picture";
        private static readonly XNamespace ContentTypeNs =
            "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly Regex ChartTagPattern = new(
            pattern: @"^Chart_(?<number>\d{3})(?<data>_data)?$",
            options: RegexOptions.CultureInvariant | RegexOptions.Compiled);
        private const string ReportDatesTag = "Report_Dates";

        private sealed record ReportTag(XElement Control, string Tag, int ChartNumber, bool IsData);
        private sealed record ReportImage(byte[] Png, int WidthMm, int HeightMm);

        #endregion

        #region Generate Report

        private async void btnReportGeneration_Click(object sender, RoutedEventArgs e)
        {
            if (!btnReportGeneration.IsEnabled)
            {
                return;
            }

            btnReportGeneration.IsEnabled = false;
            barReportGeneration.Value = 0;
            barReportGeneration.Visibility = Visibility.Visible;
            txtReportTaskComplete.Visibility = Visibility.Collapsed;
            txtReportGenerationStatus.Text = "Checking chart references...";

            try
            {
                int projectId = _activeProjectId ??
                    throw new InvalidOperationException("Select an active project.");
                DateTime start = dpReportStartDate.SelectedDate?.Date ??
                    throw new InvalidOperationException("Select Report Start Date.");
                DateTime end = dpReportEndDate.SelectedDate?.Date ??
                    throw new InvalidOperationException("Select Report End Date.");
                if (start >= end)
                {
                    throw new InvalidOperationException(
                        "Report Start Date must be earlier than Report End Date.");
                }

                if (string.IsNullOrWhiteSpace(value: txtReportName.Text))
                {
                    throw new InvalidOperationException(
                        "Enter a Report Name before generating the report.");
                }
                await SaveReportNameAsync();

                string templatePath = ValidateReportSelectionPath(
                    path: txtReportTemplatePath.Text, isTemplate: true);
                string outputFolder = ValidateReportSelectionPath(
                    path: txtReportOutputFolder.Text, isTemplate: false);
                string fileName = CreateReportFileName(
                    reportName: txtReportName.Text,
                    reportEndDate: end);
                string finalPath = Path.Combine(path1: outputFolder, path2: fileName);
                FileInfo? existingFile = File.Exists(path: finalPath)
                    ? new FileInfo(fileName: finalPath)
                    : null;
                bool overwriteExisting = existingFile is not null;
                long? existingLength = existingFile?.Length;
                DateTime? existingWriteTimeUtc = existingFile?.LastWriteTimeUtc;
                if (overwriteExisting && !ConfirmReportOverwrite(filePath: finalPath))
                {
                    barReportGeneration.Visibility = Visibility.Collapsed;
                    txtReportGenerationStatus.Text = "Report generation cancelled.";
                    return;
                }

                Dictionary<int, int> charts = await LoadReportChartIdsAsync(projectId: projectId);

                using ZipArchive template = ZipFile.OpenRead(archiveFileName: templatePath);
                ZipArchiveEntry documentEntry = template.GetEntry(entryName: "word/document.xml") ??
                    throw new InvalidDataException("The template has no Word document part.");
                XDocument document = ReadXml(entry: documentEntry);
                if (FindReportDateControls(document: document).Count != 1)
                {
                    throw new InvalidDataException(
                        "Place exactly one plain-text content control tagged Report_Dates in the front-page date cell.");
                }
                List<ReportTag> tags = FindReportTags(document: document);
                if (tags.Count == 0)
                {
                    throw new InvalidDataException(
                        "The template contains no Chart_### content-control tags.");
                }

                string[] missing = tags
                    .Where(predicate: item => !charts.ContainsKey(key: item.ChartNumber))
                    .Select(selector: item => item.Tag)
                    .Distinct(comparer: StringComparer.Ordinal)
                    .OrderBy(keySelector: item => item, comparer: StringComparer.Ordinal)
                    .ToArray();
                if (missing.Length > 0)
                {
                    MessageBoxResult decision = MessageBox.Show(
                        owner: this,
                        messageBoxText: "The following chart tags have no saved chart:\n" +
                            string.Join(separator: "\n", values: missing) +
                            "\n\nProceed with a blank 'No Chart' image for each?",
                        caption: "Missing chart references",
                        button: MessageBoxButton.YesNo,
                        icon: MessageBoxImage.Warning);
                    if (decision != MessageBoxResult.Yes)
                    {
                        barReportGeneration.Visibility = Visibility.Collapsed;
                        txtReportGenerationStatus.Text = "Report generation cancelled.";
                        return;
                    }
                }

                Dictionary<string, List<ReportImage>> images = new(StringComparer.Ordinal);
                int completed = 0;
                foreach (ReportTag tag in tags)
                {
                    if (!images.ContainsKey(key: tag.Tag))
                    {
                        images[tag.Tag] = charts.TryGetValue(key: tag.ChartNumber,
                            value: out int chartId)
                            ? await RenderSavedReportChartAsync(
                                chartId: chartId, isData: tag.IsData)
                            : [CreateNoChartImage()];
                    }

                    completed++;
                    barReportGeneration.Value = completed * 100.0 / tags.Count;
                    txtReportGenerationStatus.Text =
                        $"Prepared {completed} of {tags.Count}: {tag.Tag}";
                    await Dispatcher.Yield(priority: DispatcherPriority.Background);
                }

                string temporaryPath = Path.Combine(
                    path1: outputFolder,
                    path2: $".{Guid.NewGuid():N}.docx");
                try
                {
                    File.Copy(sourceFileName: templatePath,
                        destFileName: temporaryPath, overwrite: false);
                    WriteReportImages(
                        packagePath: temporaryPath, tags: tags, images: images,
                        reportStartDate: start, reportEndDate: end);
                    VerifyGeneratedReport(
                        packagePath: temporaryPath,
                        expectedImageCount: tags.Sum(
                            selector: tag => images[tag.Tag].Count),
                        expectedDatesText: CreateReportDatesText(
                            reportStartDate: start, reportEndDate: end));
                    if (overwriteExisting && existingFile is not null)
                    {
                        FileInfo currentFile = new(fileName: finalPath);
                        if (!currentFile.Exists ||
                            currentFile.Length != existingLength ||
                            currentFile.LastWriteTimeUtc != existingWriteTimeUtc)
                        {
                            throw new IOException(
                                "The existing report changed after overwrite was confirmed. No file was replaced.");
                        }
                    }
                    File.Move(sourceFileName: temporaryPath, destFileName: finalPath,
                        overwrite: overwriteExisting);
                }
                finally
                {
                    if (File.Exists(path: temporaryPath))
                    {
                        File.Delete(path: temporaryPath);
                    }
                }

                if (!File.Exists(path: finalPath) ||
                    new FileInfo(fileName: finalPath).Length == 0)
                {
                    throw new IOException("The completed report file is unavailable or empty.");
                }

                barReportGeneration.Visibility = Visibility.Collapsed;
                txtReportTaskComplete.Visibility = Visibility.Visible;
                txtReportGenerationStatus.Text = $"Report generated: {finalPath}";
            }
            catch (Exception ex)
            {
                barReportGeneration.Visibility = Visibility.Collapsed;
                txtReportGenerationStatus.Text = string.Empty;
                MessageBox.Show(
                    owner: this,
                    messageBoxText: ex.Message,
                    caption: "Report Generation",
                    button: MessageBoxButton.OK,
                    icon: MessageBoxImage.Error);
            }
            finally
            {
                UpdateReportGenerationAvailability();
            }
        }

        private bool ConfirmReportOverwrite(string filePath)
        {
            Window dialog = new()
            {
                Owner = this,
                Title = "Report already exists",
                Width = 540,
                MinHeight = 204,
                MinWidth = 440,
                SizeToContent = SizeToContent.Height,
                ResizeMode = ResizeMode.NoResize,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ShowInTaskbar = false
            };
            TextBlock message = new()
            {
                Text = $"The report already exists:\n{filePath}",
                TextWrapping = TextWrapping.Wrap
            };
            Button overwrite = new()
            {
                Content = "Overwrite",
                Width = 100,
                Height = 28,
                Margin = new Thickness(left: 0, top: 0, right: 8, bottom: 0)
            };
            Button cancel = new()
            {
                Content = "Cancel",
                Width = 100,
                Height = 28,
                IsCancel = true,
                IsDefault = true
            };
            overwrite.Click += (sender, e) => dialog.DialogResult = true;
            cancel.Click += (sender, e) => dialog.DialogResult = false;
            StackPanel buttons = new()
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(left: 0, top: 14, right: 0, bottom: 0)
            };
            buttons.Children.Add(element: overwrite);
            buttons.Children.Add(element: cancel);
            StackPanel content = new()
            {
                Margin = new Thickness(uniformLength: 20)
            };
            content.Children.Add(element: message);
            content.Children.Add(element: buttons);
            dialog.Content = content;
            return dialog.ShowDialog() == true;
        }

        private async Task<Dictionary<int, int>> LoadReportChartIdsAsync(int projectId)
        {
            const string sql = """
                SELECT [ChartNumber], [ChartDefinition_ID]
                FROM [dbo].[ChartDefinition]
                WHERE [Project_ID] = @Project_ID AND [IsDeleted] = 0;
                """;
            Dictionary<int, int> charts = new();
            await using SqlConnection connection = new(
                connectionString: GetTrackGeometryConnectionString());
            await connection.OpenAsync();
            await using SqlCommand command = new(cmdText: sql, connection: connection);
            command.Parameters.Add(parameterName: "@Project_ID", sqlDbType: SqlDbType.Int)
                .Value = projectId;
            await using SqlDataReader reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                charts.Add(key: reader.GetInt32(0), value: reader.GetInt32(1));
            }
            return charts;
        }

        #endregion

        #region Render Saved Chart Images

        private async Task<List<ReportImage>> RenderSavedReportChartAsync(
            int chartId, bool isData)
        {
            await LoadChartDefinitionIntoEditorAsync(chartDefinitionId: chartId);
            ApplyChartAppearanceToSeries();
            await LoadChartPreviewDataAsync(includeSeriesData: true);

            if (isData)
            {
                if (!double.TryParse(s: txtChartDataTableFontSize.Text,
                    style: NumberStyles.Float, provider: CultureInfo.InvariantCulture,
                    result: out double fontSize) || fontSize <= 0)
                {
                    throw new InvalidOperationException(
                        "Appearance: Data chart font size must be positive.");
                }

                BuildChartDataTablePreviewPages(fontSizePoints: fontSize);
                int widthMm = rbChartDataTableLandscape.IsChecked == true ? 250 : 150;
                int heightMm = GetSelectedDataChartHeightMm();
                return _chartDataTablePreviewPages
                    .Select(selector: page => new ReportImage(
                        Png: EncodePng(bitmap: page),
                        WidthMm: widthMm, HeightMm: heightMm))
                    .ToList();
            }

            if (!TryValidateChartCanvasConfiguration(
                validationMessage: out string validationMessage))
            {
                throw new InvalidOperationException(validationMessage);
            }

            int chartWidthMm = int.Parse(s: txtChartWidthMm.Text,
                provider: CultureInfo.InvariantCulture);
            int chartHeightMm = int.Parse(s: txtChartHeightMm.Text,
                provider: CultureInfo.InvariantCulture);
            int dpi = GetSelectedChartResolutionDpi();
            double widthDips = chartWidthMm * 96.0 / 25.4;
            double heightDips = chartHeightMm * 96.0 / 25.4;
            #region Capture Independent Chart Canvas

            Canvas captureCanvas = new()
            {
                Width = widthDips,
                Height = heightDips,
                Background = Brushes.White,
                ClipToBounds = true
            };
            captureCanvas.Measure(availableSize: new Size(
                width: widthDips, height: heightDips));
            captureCanvas.Arrange(finalRect: new Rect(
                x: 0, y: 0, width: widthDips, height: heightDips));
            RenderBlankChartCanvas(targetCanvas: captureCanvas);
            if (captureCanvas.Children.Count == 0)
            {
                throw new InvalidOperationException(
                    "The chart renderer produced no visual elements.");
            }
            captureCanvas.Measure(availableSize: new Size(
                width: widthDips, height: heightDips));
            captureCanvas.Arrange(finalRect: new Rect(
                x: 0, y: 0, width: widthDips, height: heightDips));
            captureCanvas.UpdateLayout();

            int pixelWidth = (int)Math.Round(
                chartWidthMm * dpi / 25.4, MidpointRounding.AwayFromZero);
            int pixelHeight = (int)Math.Round(
                chartHeightMm * dpi / 25.4, MidpointRounding.AwayFromZero);
            RenderTargetBitmap bitmap = new(
                pixelWidth: pixelWidth, pixelHeight: pixelHeight,
                dpiX: dpi, dpiY: dpi, pixelFormat: PixelFormats.Pbgra32);
            bitmap.Render(visual: captureCanvas);
            bitmap.Freeze();

            int stride = checked(pixelWidth * 4);
            byte[] pixels = new byte[checked(stride * pixelHeight)];
            bitmap.CopyPixels(pixels: pixels, stride: stride, offset: 0);
            int interiorNonWhitePixels = 0;
            int left = pixelWidth / 10;
            int right = pixelWidth - left;
            int top = pixelHeight / 10;
            int bottom = pixelHeight - top;
            for (int y = top; y < bottom; y++)
            {
                for (int x = left; x < right; x++)
                {
                    int offset = checked(y * stride + x * 4);
                    if (pixels[offset] != 255 ||
                        pixels[offset + 1] != 255 ||
                        pixels[offset + 2] != 255)
                    {
                        interiorNonWhitePixels++;
                    }
                }
            }
            int interiorPixelCount = checked((right - left) * (bottom - top));
            if (interiorNonWhitePixels < interiorPixelCount / 1000)
            {
                throw new InvalidOperationException(
                    "Chart image capture contains no visible plot-area content. " +
                    "Report generation was stopped to avoid saving a blank chart.");
            }

            return [new ReportImage(
                Png: EncodePng(bitmap: bitmap),
                WidthMm: chartWidthMm, HeightMm: chartHeightMm)];

            #endregion
        }

        private static byte[] EncodePng(BitmapSource bitmap)
        {
            PngBitmapEncoder encoder = new();
            encoder.Frames.Add(item: BitmapFrame.Create(source: bitmap));
            using MemoryStream stream = new();
            encoder.Save(stream: stream);
            return stream.ToArray();
        }

        private static ReportImage CreateNoChartImage()
        {
            const int widthMm = 150;
            const int heightMm = 80;
            const int dpi = 300;
            double widthDips = widthMm * 96.0 / 25.4;
            double heightDips = heightMm * 96.0 / 25.4;
            DrawingVisual visual = new();
            using (DrawingContext context = visual.RenderOpen())
            {
                context.DrawRectangle(brush: Brushes.White, pen: null,
                    rectangle: new Rect(x: 0, y: 0,
                        width: widthDips, height: heightDips));
                FormattedText label = new(
                    textToFormat: "No Chart", culture: CultureInfo.InvariantCulture,
                    flowDirection: FlowDirection.LeftToRight,
                    typeface: new Typeface(typefaceName: "Arial"),
                    emSize: 20.0 * 96.0 / 72.0,
                    foreground: Brushes.LightGray, pixelsPerDip: 1.0);
                context.DrawText(formattedText: label,
                    origin: new Point(
                        x: (widthDips - label.Width) / 2.0,
                        y: (heightDips - label.Height) / 2.0));
            }
            RenderTargetBitmap bitmap = new(
                pixelWidth: (int)Math.Round(widthMm * dpi / 25.4,
                    MidpointRounding.AwayFromZero),
                pixelHeight: (int)Math.Round(heightMm * dpi / 25.4,
                    MidpointRounding.AwayFromZero),
                dpiX: dpi, dpiY: dpi, pixelFormat: PixelFormats.Pbgra32);
            bitmap.Render(visual: visual);
            bitmap.Freeze();
            return new ReportImage(Png: EncodePng(bitmap: bitmap),
                WidthMm: widthMm, HeightMm: heightMm);
        }

        #endregion

        #region Read Tagged Word Cells

        private static XDocument ReadXml(ZipArchiveEntry entry)
        {
            using Stream stream = entry.Open();
            return XDocument.Load(stream: stream);
        }

        private static List<ReportTag> FindReportTags(XDocument document)
        {
            List<ReportTag> tags = new();
            foreach (XElement control in document.Descendants(name: WordNs + "sdt"))
            {
                string? value = (string?)control.Element(name: WordNs + "sdtPr")?
                    .Element(name: WordNs + "tag")?.Attribute(name: WordNs + "val");
                if (value is null)
                {
                    continue;
                }
                Match match = ChartTagPattern.Match(input: value);
                if (!match.Success)
                {
                    continue;
                }
                tags.Add(item: new ReportTag(
                    Control: control, Tag: value,
                    ChartNumber: int.Parse(s: match.Groups["number"].Value,
                        provider: CultureInfo.InvariantCulture),
                    IsData: match.Groups["data"].Success));
            }
            return tags;
        }

        private static List<XElement> FindReportDateControls(XDocument document)
        {
            return document.Descendants(name: WordNs + "sdt")
                .Where(predicate: control => string.Equals(
                    a: (string?)control.Element(name: WordNs + "sdtPr")?
                        .Element(name: WordNs + "tag")?
                        .Attribute(name: WordNs + "val"),
                    b: ReportDatesTag,
                    comparisonType: StringComparison.Ordinal))
                .ToList();
        }

        private static string CreateReportDatesText(
            DateTime reportStartDate, DateTime reportEndDate)
        {
            return $"Report period: {reportStartDate:yyyy.MM.dd} to {reportEndDate:yyyy.MM.dd}";
        }

        private static void SetReportDatesText(
            XDocument document, DateTime reportStartDate, DateTime reportEndDate)
        {
            XElement control = FindReportDateControls(document: document).Single();
            XElement content = control.Element(name: WordNs + "sdtContent") ??
                throw new InvalidDataException("Report_Dates has no content.");
            XElement run = new(WordNs + "r",
                new XElement(WordNs + "rPr",
                    new XElement(WordNs + "rFonts",
                        new XAttribute(WordNs + "ascii", "Arial"),
                        new XAttribute(WordNs + "hAnsi", "Arial")),
                    new XElement(WordNs + "b"),
                    new XElement(WordNs + "sz",
                        new XAttribute(WordNs + "val", 24)),
                    new XElement(WordNs + "szCs",
                        new XAttribute(WordNs + "val", 24))),
                new XElement(WordNs + "t",
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    CreateReportDatesText(
                        reportStartDate: reportStartDate,
                        reportEndDate: reportEndDate)));

            if (control.Parent?.Name == WordNs + "p")
            {
                XElement paragraph = control.Parent;
                XElement properties = paragraph.Element(name: WordNs + "pPr") ??
                    new XElement(WordNs + "pPr");
                properties.Element(name: WordNs + "jc")?.Remove();
                properties.Add(content: new XElement(WordNs + "jc",
                    new XAttribute(WordNs + "val", "center")));
                if (properties.Parent is null)
                {
                    paragraph.AddFirst(content: properties);
                }
                content.ReplaceNodes(content: run);
            }
            else if (control.Parent?.Name == WordNs + "tr")
            {
                XElement cell = content.Element(name: WordNs + "tc") ??
                    throw new InvalidDataException(
                        "Report_Dates must retain its table cell within the row.");
                cell.Elements().Where(predicate: element =>
                    element.Name != WordNs + "tcPr").Remove();
                cell.Add(content: new XElement(WordNs + "p",
                    new XElement(WordNs + "pPr",
                        new XElement(WordNs + "jc",
                            new XAttribute(WordNs + "val", "center"))),
                    run));
            }
            else
            {
                content.ReplaceNodes(content: new XElement(WordNs + "p",
                    new XElement(WordNs + "pPr",
                        new XElement(WordNs + "jc",
                            new XAttribute(WordNs + "val", "center"))),
                    run));
            }
        }

        #endregion

        #region Insert Images Into Word Package

        private static void WriteReportImages(string packagePath,
            IReadOnlyList<ReportTag> tags,
            IReadOnlyDictionary<string, List<ReportImage>> images,
            DateTime reportStartDate, DateTime reportEndDate)
        {
            using ZipArchive package = ZipFile.Open(
                archiveFileName: packagePath, mode: ZipArchiveMode.Update);
            ZipArchiveEntry documentEntry = package.GetEntry(entryName: "word/document.xml") ??
                throw new InvalidDataException("The Word document part is missing.");
            XDocument document = ReadXml(entry: documentEntry);
            List<ReportTag> outputTags = FindReportTags(document: document);
            if (outputTags.Count != tags.Count)
            {
                throw new InvalidDataException("The template tags changed while generating the report.");
            }
            SetReportDatesText(document: document,
                reportStartDate: reportStartDate,
                reportEndDate: reportEndDate);

            ZipArchiveEntry? relationshipEntry =
                package.GetEntry(entryName: "word/_rels/document.xml.rels");
            XDocument relationships = relationshipEntry is null
                ? new XDocument(new XElement(PackageRelationshipNs + "Relationships"))
                : ReadXml(entry: relationshipEntry);
            XElement relationshipRoot = relationships.Root ??
                throw new InvalidDataException("The relationship part is invalid.");
            int nextRelationshipId = 1;
            int nextPictureId = document.Descendants(name: WordDrawingNs + "docPr")
                .Select(selector: item => int.TryParse(
                    s: (string?)item.Attribute(name: "id"),
                    result: out int existingId) ? existingId : 0)
                .DefaultIfEmpty(defaultValue: 0)
                .Max() + 1;
            foreach (ReportTag tag in outputTags)
            {
                List<ReportImage> pages = images[tag.Tag];
                XElement content = tag.Control.Element(name: WordNs + "sdtContent") ??
                    throw new InvalidDataException($"Tag {tag.Tag} has no content cell.");
                XElement cell = content.Element(name: WordNs + "tc") ??
                    throw new InvalidDataException(
                        $"Tag {tag.Tag} must enclose a complete table cell.");
                cell.Elements().Where(predicate: element =>
                    element.Name != WordNs + "tcPr").Remove();

                foreach (ReportImage page in pages)
                {
                    string relationshipId;
                    do
                    {
                        relationshipId = $"rIdGnaReport{nextRelationshipId++}";
                    }
                    while (relationshipRoot.Elements().Any(predicate: item =>
                        (string?)item.Attribute(name: "Id") == relationshipId));

                    string imageName = $"gna-report-{Guid.NewGuid():N}.png";
                    string imagePath = "word/media/" + imageName;
                    ZipArchiveEntry imageEntry = package.CreateEntry(
                        entryName: imagePath, compressionLevel: CompressionLevel.Optimal);
                    using (Stream imageStream = imageEntry.Open())
                    {
                        imageStream.Write(buffer: page.Png);
                    }
                    relationshipRoot.Add(new XElement(
                        PackageRelationshipNs + "Relationship",
                        new XAttribute("Id", relationshipId),
                        new XAttribute("Type",
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
                        new XAttribute("Target", "media/" + imageName)));
                    cell.Add(content: CreateImageParagraph(
                        relationshipId: relationshipId,
                        pictureId: nextPictureId++,
                        imageName: imageName,
                        widthMm: page.WidthMm,
                        heightMm: page.HeightMm));
                }
            }

            ZipArchiveEntry contentTypeEntry = package.GetEntry(
                entryName: "[Content_Types].xml") ??
                throw new InvalidDataException("The package content types are missing.");
            XDocument contentTypes = ReadXml(entry: contentTypeEntry);
            XElement typeRoot = contentTypes.Root ??
                throw new InvalidDataException("The package content types are invalid.");
            if (!typeRoot.Elements().Any(predicate: element =>
                (string?)element.Attribute(name: "Extension") == "png"))
            {
                typeRoot.Add(content: new XElement(
                    ContentTypeNs + "Default",
                    new XAttribute("Extension", "png"),
                    new XAttribute("ContentType", "image/png")));
            }

            ReplaceXml(package: package, entryName: "word/document.xml", document: document);
            ReplaceXml(package: package, entryName: "word/_rels/document.xml.rels",
                document: relationships);
            ReplaceXml(package: package, entryName: "[Content_Types].xml",
                document: contentTypes);
        }

        private static void VerifyGeneratedReport(
            string packagePath, int expectedImageCount,
            string expectedDatesText)
        {
            using ZipArchive package = ZipFile.OpenRead(archiveFileName: packagePath);
            XDocument document = ReadXml(entry: package.GetEntry(
                entryName: "word/document.xml") ??
                throw new InvalidDataException("The generated document part is missing."));
            if (document.Descendants(name: WordNs + "sdt")
                .Any(predicate: control =>
                    control.Parent?.Name == WordNs + "tr" &&
                    control.Element(name: WordNs + "sdtContent")?
                        .Element(name: WordNs + "tc") is null))
            {
                throw new InvalidDataException(
                    "A tagged table row lost its table cell during report generation.");
            }
            if (!document.Descendants(name: WordNs + "t")
                .Any(predicate: text => text.Value == expectedDatesText))
            {
                throw new InvalidDataException(
                    "The report start and end dates are missing from the document.");
            }
            XElement dateControl = FindReportDateControls(document: document).Single();
            if (dateControl.Parent?.Name == WordNs + "tr" &&
                dateControl.Element(name: WordNs + "sdtContent")?
                    .Elements(name: WordNs + "tc").Count() != 1)
            {
                throw new InvalidDataException(
                    "The report-date content control no longer contains a table cell.");
            }
            XDocument relationships = ReadXml(entry: package.GetEntry(
                entryName: "word/_rels/document.xml.rels") ??
                throw new InvalidDataException("The generated image relationships are missing."));
            _ = ReadXml(entry: package.GetEntry(
                entryName: "[Content_Types].xml") ??
                throw new InvalidDataException("The generated content types are missing."));

            Dictionary<string, string> generatedImages = relationships.Root?
                .Elements(name: PackageRelationshipNs + "Relationship")
                .Where(predicate: item =>
                    ((string?)item.Attribute(name: "Target"))?
                        .StartsWith(value: "media/gna-report-",
                            comparisonType: StringComparison.Ordinal) == true)
                .ToDictionary(
                    keySelector: item => (string?)item.Attribute(name: "Id") ?? string.Empty,
                    elementSelector: item => (string?)item.Attribute(name: "Target") ?? string.Empty,
                    comparer: StringComparer.Ordinal)
                ?? throw new InvalidDataException("The relationship part is empty.");

            if (generatedImages.Count != expectedImageCount)
            {
                throw new InvalidDataException(
                    "The report does not contain the expected number of generated images.");
            }

            HashSet<string> embedded = document
                .Descendants(name: DrawingNs + "blip")
                .Select(selector: element =>
                    (string?)element.Attribute(name: RelationshipNs + "embed") ?? string.Empty)
                .ToHashSet(comparer: StringComparer.Ordinal);
            byte[] pngSignature = [137, 80, 78, 71, 13, 10, 26, 10];
            foreach ((string id, string target) in generatedImages)
            {
                if (!embedded.Contains(item: id))
                {
                    throw new InvalidDataException(
                        $"Generated image relationship {id} is not embedded in the document.");
                }

                ZipArchiveEntry entry = package.GetEntry(
                    entryName: "word/" + target) ??
                    throw new InvalidDataException($"Generated image {target} is missing.");
                if (entry.Length <= pngSignature.Length)
                {
                    throw new InvalidDataException($"Generated image {target} is empty.");
                }
                using Stream image = entry.Open();
                byte[] signature = new byte[pngSignature.Length];
                image.ReadExactly(buffer: signature);
                if (!signature.SequenceEqual(second: pngSignature))
                {
                    throw new InvalidDataException($"Generated image {target} is not a PNG.");
                }
            }
        }

        private static void ReplaceXml(ZipArchive package,
            string entryName, XDocument document)
        {
            package.GetEntry(entryName: entryName)?.Delete();
            ZipArchiveEntry entry = package.CreateEntry(
                entryName: entryName, compressionLevel: CompressionLevel.Optimal);
            using Stream stream = entry.Open();
            document.Save(stream: stream);
        }

        private static XElement CreateImageParagraph(string relationshipId,
            int pictureId, string imageName, int widthMm, int heightMm)
        {
            long cx = widthMm * 36000L;
            long cy = heightMm * 36000L;
            XElement properties = new(PictureNs + "nvPicPr",
                new XElement(PictureNs + "cNvPr",
                    new XAttribute("id", pictureId),
                    new XAttribute("name", imageName)),
                new XElement(PictureNs + "cNvPicPr"));
            XElement fill = new(PictureNs + "blipFill",
                new XElement(DrawingNs + "blip",
                    new XAttribute(RelationshipNs + "embed", relationshipId)),
                new XElement(DrawingNs + "stretch",
                    new XElement(DrawingNs + "fillRect")));
            XElement shape = new(PictureNs + "spPr",
                new XElement(DrawingNs + "xfrm",
                    new XElement(DrawingNs + "off",
                        new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(DrawingNs + "ext",
                        new XAttribute("cx", cx), new XAttribute("cy", cy))),
                new XElement(DrawingNs + "prstGeom",
                    new XAttribute("prst", "rect"),
                    new XElement(DrawingNs + "avLst")));
            XElement picture = new(PictureNs + "pic", properties, fill, shape);
            XElement graphic = new(DrawingNs + "graphic",
                new XElement(DrawingNs + "graphicData",
                    new XAttribute("uri",
                        "http://schemas.openxmlformats.org/drawingml/2006/picture"),
                    picture));
            XElement inline = new(WordDrawingNs + "inline",
                new XAttribute("distT", 0), new XAttribute("distB", 0),
                new XAttribute("distL", 0), new XAttribute("distR", 0),
                new XElement(WordDrawingNs + "extent",
                    new XAttribute("cx", cx), new XAttribute("cy", cy)),
                new XElement(WordDrawingNs + "docPr",
                    new XAttribute("id", pictureId),
                    new XAttribute("name", imageName)),
                graphic);
            return new XElement(WordNs + "p",
                new XElement(WordNs + "r",
                    new XElement(WordNs + "drawing", inline)));
        }

        #endregion
    }
}
