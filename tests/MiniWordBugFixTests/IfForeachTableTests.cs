using MiniSoftware;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using Xunit.Abstractions;

namespace MiniWordBugFixTests;

/// <summary>
/// Tests for MiniWord bugs that need fixing:
/// Bug 1: @if varName (truthy-only) negates the value
/// Bug 2: @if inside @foreach crashes
/// Bug 3: @if inside TableStart/TableEnd
/// </summary>
public class IfForeachTableTests
{
    private readonly ITestOutputHelper _output;
    private static readonly string OutputDir = Path.Combine(Path.GetTempPath(), "MiniWordBugFixTests");

    public IfForeachTableTests(ITestOutputHelper output)
    {
        _output = output;
        Directory.CreateDirectory(OutputDir);
    }

    #region Bug 1: @if varName negates value

    [Fact]
    public void If_Truthy_True_ShouldShowContent()
    {
        var path = Tpl(
            "@if is_vip",
            "VIP content",
            "@endif"
        );

        var outPath = Out("if_truthy_true.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["is_vip"] = "true"
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("VIP content", content);
    }

    [Fact]
    public void If_Truthy_False_ShouldHideContent()
    {
        var path = Tpl(
            "@if is_vip",
            "VIP content",
            "@endif"
        );

        var outPath = Out("if_truthy_false.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["is_vip"] = "false"
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.DoesNotContain("VIP content", content);
    }

    #endregion

    #region Bug 2: @if inside @foreach

    [Fact]
    public void IfInsideForeach_Comparison_ShouldWork()
    {
        var path = Tpl(
            "@foreach {{items}}",
            "{{procedure}}: {{result}}",
            "@if is_abnormal == true",
            "ABNORMAL",
            "@endif",
            "@endforeach"
        );

        var outPath = Out("if_inside_foreach_comparison.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["items"] = new List<Dictionary<string, object>>
            {
                new() { ["procedure"] = "CT Head", ["result"] = "Normal", ["is_abnormal"] = "false" },
                new() { ["procedure"] = "CT Neck", ["result"] = "Abnormal", ["is_abnormal"] = "true" },
            }
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("CT Head", content);
        Assert.Contains("CT Neck", content);
        Assert.Contains("Normal", content);
        Assert.Contains("ABNORMAL", content);
        Assert.DoesNotContain("@if", content);
        Assert.DoesNotContain("@endif", content);
        Assert.DoesNotContain("@foreach", content);
        Assert.DoesNotContain("@endforeach", content);
    }

    [Fact]
    public void IfInsideForeach_ConditionalFlag_ShouldWork()
    {
        var path = Tpl(
            "@foreach {{items}}",
            "{{procedure}}: {{result}} {{abnormal_flag}}",
            "@endif",
            "@endforeach"
        );

        // Actually this test uses @if, not just @endif
        // Let me restructure:
        // Template: @foreach + @if for conditional display
        var path2 = Tpl(
            "@foreach {{items}}",
            "{{procedure}}: {{result}}",
            "@if is_abnormal == true",
            "[ABNORMAL]",
            "@endif",
            "@endforeach"
        );

        var outPath = Out("if_inside_foreach_flag.docx");
        MiniWord.SaveAsByTemplate(outPath, path2, new Dictionary<string, object>
        {
            ["items"] = new List<Dictionary<string, object>>
            {
                new() { ["procedure"] = "CT Head", ["result"] = "Normal", ["is_abnormal"] = "false" },
                new() { ["procedure"] = "CT Neck", ["result"] = "Abnormal", ["is_abnormal"] = "true" },
            }
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("[ABNORMAL]", content);
    }

    [Fact]
    public void IfInsideForeach_AllItemsConditional()
    {
        var path = Tpl(
            "@foreach {{items}}",
            "@if status == active",
            "{{name}} ✓",
            "@endif",
            "@endforeach"
        );

        var outPath = Out("if_inside_foreach_all.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["items"] = new List<Dictionary<string, object>>
            {
                new() { ["name"] = "Item A", ["status"] = "active" },
                new() { ["name"] = "Item B", ["status"] = "inactive" },
                new() { ["name"] = "Item C", ["status"] = "active" },
            }
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("Item A", content);
        Assert.Contains("Item C", content);
        Assert.DoesNotContain("Item B", content);
    }

    #endregion

    #region Bug 3: @if inside TableStart/TableEnd

    [Fact(Skip = "TableStart/TableEnd requires Word-created template; tested separately")]
    public void IfInsideTable_Comparison_ShouldWork()
    {
        var path = CreateTableTemplate();

        var outPath = Out("if_inside_table.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["items"] = new List<Dictionary<string, object>>
            {
                new() { ["procedure"] = "CT Head", ["result"] = "Normal", ["is_abnormal"] = "false" },
                new() { ["procedure"] = "CT Neck", ["result"] = "Abnormal", ["is_abnormal"] = "true" },
            }
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("CT Head", content);
        Assert.Contains("CT Neck", content);
    }

    #endregion

    #region Regression: existing features still work

    [Fact]
    public void If_Comparison_StringEq_ShouldWork()
    {
        var path = Tpl(
            "@if status == active",
            "Active content",
            "@endif"
        );

        var outPath = Out("if_comparison_eq.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["status"] = "active"
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("Active content", content);
        Assert.DoesNotContain("@if", content);
    }

    [Fact]
    public void If_Comparison_StringNeq_ShouldHide()
    {
        var path = Tpl(
            "@if status == active",
            "Active content",
            "@endif"
        );

        var outPath = Out("if_comparison_neq.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["status"] = "inactive"
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.DoesNotContain("Active content", content);
    }

    [Fact]
    public void If_Comparison_Numeric_ShouldWork()
    {
        var path = Tpl(
            "@if age > 18",
            "Adult",
            "@endif"
        );

        var outPath = Out("if_numeric.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["age"] = 25
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("Adult", content);
    }

    [Fact]
    public void Foreach_Basic_ShouldWork()
    {
        var path = Tpl(
            "@foreach {{items}}",
            "Item: {{name}}",
            "@endforeach"
        );

        var outPath = Out("foreach_basic.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["items"] = new List<Dictionary<string, object>>
            {
                new() { ["name"] = "Alpha" },
                new() { ["name"] = "Beta" },
            }
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("Alpha", content);
        Assert.Contains("Beta", content);
    }

    [Fact]
    public void Foreach_EmptyList_ShouldRemoveBlock()
    {
        var path = Tpl(
            "Before",
            "@foreach {{items}}",
            "Item: {{name}}",
            "@endforeach",
            "After"
        );

        var outPath = Out("foreach_empty.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["items"] = new List<Dictionary<string, object>>()
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("Before", content);
        Assert.Contains("After", content);
        Assert.DoesNotContain("@foreach", content);
    }

    [Fact]
    public void If_Truthy_NestedInsideForeach()
    {
        var path = Tpl(
            "@foreach {{items}}",
            "Name: {{name}}",
            "@if is_vip",
            " [VIP]",
            "@endif",
            "@endforeach"
        );

        var outPath = Out("if_truthy_in_foreach.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["items"] = new List<Dictionary<string, object>>
            {
                new() { ["name"] = "Alice", ["is_vip"] = "true" },
                new() { ["name"] = "Bob", ["is_vip"] = "false" },
            }
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("Alice", content);
        Assert.Contains("[VIP]", content);
        Assert.Contains("Bob", content);
    }

    [Fact]
    public void IfInsideForeach_AllFalse()
    {
        var path = Tpl(
            "@foreach {{items}}",
            "Name: {{name}}",
            "@if active == true",
            " [ACTIVE]",
            "@endif",
            "@endforeach"
        );

        var outPath = Out("if_inside_foreach_all_false.docx");
        MiniWord.SaveAsByTemplate(outPath, path, new Dictionary<string, object>
        {
            ["items"] = new List<Dictionary<string, object>>
            {
                new() { ["name"] = "Item1", ["active"] = "false" },
                new() { ["name"] = "Item2", ["active"] = "false" },
            }
        });

        var content = Read(outPath);
        _output.WriteLine(content);
        Assert.Contains("Item1", content);
        Assert.Contains("Item2", content);
        Assert.DoesNotContain("[ACTIVE]", content);
    }

    #endregion

    #region Helpers

    private static string Tpl(params string[] lines)
    {
        var hash = string.Join("|", lines);
        var path = Path.Combine(OutputDir, $"tpl_{Math.Abs(hash.GetHashCode()):x}.docx");
        if (File.Exists(path)) return path;

        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            foreach (var line in lines)
                body.Append(P(line));
            body.Append(new SectionProperties());
            mainPart.Document.Append(body);
            mainPart.Document.Save();
        }
        File.WriteAllBytes(path, ms.ToArray());
        return path;
    }

    private static string CreateTableTemplate()
    {
        var path = Path.Combine(OutputDir, "tpl_if_table.docx");

        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();

            var table = new Table();
            var props = new TableProperties
            {
                TableBorders = new TableBorders
                {
                    TopBorder = new TopBorder { Val = BorderValues.Single, Size = 1 },
                    BottomBorder = new BottomBorder { Val = BorderValues.Single, Size = 1 },
                    LeftBorder = new LeftBorder { Val = BorderValues.Single, Size = 1 },
                    RightBorder = new RightBorder { Val = BorderValues.Single, Size = 1 },
                }
            };
            table.Append(props);

            // Header row
            var headerRow = new TableRow();
            headerRow.Append(Cell("Procedure"));
            headerRow.Append(Cell("Result"));
            table.Append(headerRow);

            // Data row with TableStart/TableEnd
            var dataRow = new TableRow();
            var cell1 = new TableCell();
            cell1.Append(P("{{TableStart:items}}{{items.procedure}}"));
            var cell2 = new TableCell();
            cell2.Append(P("{{items.result}}{{TableEnd:items}}"));
            dataRow.Append(cell1);
            dataRow.Append(cell2);
            table.Append(dataRow);

            body.Append(table);
            body.Append(new SectionProperties());
            mainPart.Document.Append(body);
            mainPart.Document.Save();
        }
        File.WriteAllBytes(path, ms.ToArray());
        return path;
    }

    private static string Out(string name) => Path.Combine(OutputDir, name);

    private static Paragraph P(string text)
    {
        var p = new Paragraph();
        var r = new Run();
        r.AppendChild(new RunProperties());
        r.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        p.AppendChild(r);
        return p;
    }

    private static TableCell Cell(string text)
    {
        var cell = new TableCell();
        cell.Append(P(text));
        return cell;
    }

    private static string Read(string path)
    {
        if (!File.Exists(path)) return "(file not found)";
        using var doc = WordprocessingDocument.Open(path, false);
        return doc.MainDocumentPart?.Document?.Body?.InnerText ?? "";
    }

    #endregion
}