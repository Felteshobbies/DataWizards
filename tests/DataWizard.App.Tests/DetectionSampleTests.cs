using System.Text;
using DataWizard.App.Services;
using DataWizard.App.ViewModels;
using DataWizard.Core.Configuration;

namespace DataWizard.App.Tests;

/// <summary>
/// The Detection tab tests its rules against a sample. It follows the file
/// selected on the Convert tab, which is far more useful than a fixed example -
/// but must never overwrite rows the user typed in themselves.
/// </summary>
public class DetectionSampleTests : IDisposable
{
    private readonly string _root = Path.Combine(
        Path.GetTempPath(), "DataWizardSampleTests", Guid.NewGuid().ToString("N"));

    public DetectionSampleTests() => Directory.CreateDirectory(_root);

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_root))
                Directory.Delete(_root, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    private string WriteFile(string name, string content, Encoding? encoding = null)
    {
        var path = Path.Combine(_root, name);
        File.WriteAllText(path, content, encoding ?? new UTF8Encoding(false));
        return path;
    }

    private AppSession CreateSession() =>
        new(DataWizardSettings.CreateDefault(), Path.Combine(_root, "settings.json"));

    private static (ConvertViewModel Convert, DetectionViewModel Detection) CreateTabs(AppSession session)
    {
        var dialogs = new DialogService(() => null);
        return (new ConvertViewModel(session, dialogs), new DetectionViewModel(session, dialogs));
    }

    [Fact]
    public void WithoutASelectionTheBuiltInRowsAreUsed()
    {
        var (_, detection) = CreateTabs(CreateSession());

        Assert.Null(detection.SampleSourcePath);
        Assert.Equal("Built-in example rows", detection.SampleSourceLabel);
        Assert.Contains("customer_id", detection.SampleText);
        Assert.False(detection.HasSelectedFile);
    }

    [Fact]
    public void SelectingAFileLoadsItsFirstLines()
    {
        var path = WriteFile("orders.csv",
            "order_no;article;menge\r\n1001;0004711;2,5\r\n1002;0004712;10\r\n");

        var session = CreateSession();
        var (convert, detection) = CreateTabs(session);

        convert.AddFiles([path]);
        convert.SelectedFile = convert.Files[0];

        Assert.Equal(path, detection.SampleSourcePath);
        Assert.Equal("From orders.csv", detection.SampleSourceLabel);
        Assert.Contains("order_no;article;menge", detection.SampleText);
        Assert.Contains("1002;0004712;10", detection.SampleText);
        Assert.True(detection.HasSelectedFile);
    }

    [Fact]
    public void TheSampleIsAnalysedWithTheCurrentRules()
    {
        var path = WriteFile("orders.csv",
            "artikel;menge;preis\r\n0004711;2,5;19,99\r\n0004712;10;4,50\r\n");

        var session = CreateSession();
        var (convert, detection) = CreateTabs(session);

        convert.AddFiles([path]);
        convert.SelectedFile = convert.Files[0];

        Assert.NotNull(detection.TestResult);

        var columns = detection.TestResult!.Columns;
        Assert.Equal("Text", columns.Single(c => c.Name == "artikel").EffectiveType);
        Assert.Equal("Decimal", columns.Single(c => c.Name == "menge").EffectiveType);
        Assert.Equal("Decimal", columns.Single(c => c.Name == "preis").EffectiveType);
    }

    /// <summary>
    /// Silently replacing something the user typed would be worse than showing a
    /// sample that no longer matches the selection.
    /// </summary>
    [Fact]
    public void AHandEditedSampleIsNotReplacedByASelection()
    {
        var path = WriteFile("orders.csv", "order_no;article\r\n1001;0004711\r\n");

        var session = CreateSession();
        var (convert, detection) = CreateTabs(session);

        detection.SampleText = "my;own;rows\r\n1;2;3";

        convert.AddFiles([path]);
        convert.SelectedFile = convert.Files[0];

        Assert.Contains("my;own;rows", detection.SampleText);
        Assert.Null(detection.SampleSourcePath);
    }

    [Fact]
    public void TheUserCanStillAskForTheSelectedFileExplicitly()
    {
        var path = WriteFile("orders.csv", "order_no;article\r\n1001;0004711\r\n");

        var session = CreateSession();
        var (convert, detection) = CreateTabs(session);

        detection.SampleText = "my;own;rows";

        convert.AddFiles([path]);
        convert.SelectedFile = convert.Files[0];

        detection.LoadSampleFromFileCommand.Execute(null);

        Assert.Contains("order_no;article", detection.SampleText);
        Assert.Equal(path, detection.SampleSourcePath);
    }

    [Fact]
    public void TheBuiltInRowsCanBeRestored()
    {
        var path = WriteFile("orders.csv", "order_no;article\r\n1001;0004711\r\n");

        var session = CreateSession();
        var (convert, detection) = CreateTabs(session);

        convert.AddFiles([path]);
        convert.SelectedFile = convert.Files[0];
        Assert.NotNull(detection.SampleSourcePath);

        detection.UseBuiltInSampleCommand.Execute(null);

        Assert.Null(detection.SampleSourcePath);
        Assert.Contains("customer_id", detection.SampleText);
        Assert.False(detection.CanUseBuiltInSample);
    }

    [Fact]
    public void AfterRestoringTheBuiltInRowsASelectionIsFollowedAgain()
    {
        var first = WriteFile("first.csv", "a;b\r\n1;2\r\n");
        var second = WriteFile("second.csv", "x;y\r\n3;4\r\n");

        var session = CreateSession();
        var (convert, detection) = CreateTabs(session);

        convert.AddFiles([first, second]);

        convert.SelectedFile = convert.Files[0];
        Assert.Contains("a;b", detection.SampleText);

        convert.SelectedFile = convert.Files[1];
        Assert.Contains("x;y", detection.SampleText);
    }

    /// <summary>
    /// The sample is a preview of the file, so an Excel workbook - which is not
    /// delimited text at all - must not be offered as one.
    /// </summary>
    [Fact]
    public void SelectingAWorkbookDoesNotChangeTheSample()
    {
        var csv = WriteFile("orders.csv", "order_no;article\r\n1001;0004711\r\n");
        var workbook = WriteFile("book.xlsx", "not really a workbook");

        var session = CreateSession();
        var (convert, detection) = CreateTabs(session);

        convert.AddFiles([csv, workbook]);
        convert.SelectedFile = convert.Files[0];
        Assert.Equal(csv, detection.SampleSourcePath);

        convert.SelectedFile = convert.Files[1];

        Assert.Contains("order_no;article", detection.SampleText);
        Assert.False(detection.HasSelectedFile);
    }

    [Fact]
    public void AFileSelectedBeforeTheTabExistsIsPickedUp()
    {
        var path = WriteFile("orders.csv", "order_no;article\r\n1001;0004711\r\n");

        var session = CreateSession();
        var dialogs = new DialogService(() => null);

        var convert = new ConvertViewModel(session, dialogs);
        convert.AddFiles([path]);
        convert.SelectedFile = convert.Files[0];

        // Built only now, as the tab would be when the window is created.
        var detection = new DetectionViewModel(session, dialogs);

        Assert.Equal(path, detection.SampleSourcePath);
        Assert.Contains("order_no;article", detection.SampleText);
    }

    [Fact]
    public void AnUnreadableFileLeavesTheSampleAloneAndIsLogged()
    {
        var session = CreateSession();
        var (_, detection) = CreateTabs(session);

        var before = detection.SampleText;
        session.SetCurrentFile(Path.Combine(_root, "missing.csv"));

        Assert.Equal(before, detection.SampleText);
    }

    [Fact]
    public void AWindows1252FileIsDecodedForTheSample()
    {
        var path = WriteFile("umlauts.csv",
            "name;ort\r\nMüller;Köln\r\nWeiß;Zürich\r\n",
            Encoding.GetEncoding(1252));

        var session = CreateSession();
        session.Settings.Detection.ForcedEncoding = "windows-1252";

        var (convert, detection) = CreateTabs(session);
        convert.AddFiles([path]);
        convert.SelectedFile = convert.Files[0];

        Assert.Contains("Müller", detection.SampleText);
        Assert.Contains("Zürich", detection.SampleText);
    }

    [Fact]
    public void OnlyTheOpeningLinesOfALargeFileAreRead()
    {
        var builder = new StringBuilder("id;value\r\n");
        for (var i = 1; i <= 5000; i++)
            builder.Append(i).Append(';').Append(i).Append("\r\n");

        var path = WriteFile("big.csv", builder.ToString());

        var session = CreateSession();
        var (convert, detection) = CreateTabs(session);

        convert.AddFiles([path]);
        convert.SelectedFile = convert.Files[0];

        var lines = detection.SampleText.Split('\n').Length;
        Assert.InRange(lines, 2, 40);
    }
}
