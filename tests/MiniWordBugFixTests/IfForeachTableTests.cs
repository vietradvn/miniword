using MiniSoftware;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Compression;
using System.Text;
using System.Diagnostics;
using Xunit;
using Xunit.Abstractions;

namespace MiniWordBugFixTests;

public class IfForeachTableTests
{
    private readonly ITestOutputHelper _output;
    private static readonly string OutputDir = Path.Combine(Path.GetTempPath(), "MiniWordBugFixTests");
    private const string ValidateDir = "/tmp/docx_validate";
    private const string DockerContainer = "eis-print-dev";

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
            "{{procedure}}: {{result}}",
            "@if is_abnormal == true",
            "[ABNORMAL]",
            "@endif",
            "@endforeach"
        );

        var outPath = Out("if_inside_foreach_flag.docx");
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
        Assert.Contains("[ABNORMAL]", content);
    }

    [Fact]
    public void IfInsideForeach_AllItemsConditional()
    {
        var path = Tpl(
            "@foreach {{items}}",
            "@if status == active",
            "{{name}} checkmark",
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

    #region LibreOffice validation — ensures DOCX opens in real office software

    [Fact]
    public void LibreOffice_CanOpen_AllOutputs()
    {
        if (!DockerAvailable())
        {
            _output.WriteLine("Docker not available, skipping LibreOffice validation");
            return;
        }

        var filesToValidate = new List<(string file, string expectedContent)>();

        // Generate if_truthy_true
        {
            var tpl = Tpl("@if is_vip", "VIP content", "@endif");
            var outPath = Out("lo_if_truthy_true.docx");
            MiniWord.SaveAsByTemplate(outPath, tpl, new Dictionary<string, object> { ["is_vip"] = "true" });
            filesToValidate.Add((outPath, "VIP content"));
        }

        // Generate if_inside_foreach_comparison
        {
            var tpl = Tpl("@foreach {{items}}", "{{name}}", "@if active == true", "[ACTIVE]", "@endif", "@endforeach");
            var outPath = Out("lo_if_inside_foreach.docx");
            MiniWord.SaveAsByTemplate(outPath, tpl, new Dictionary<string, object>
            {
                ["items"] = new List<Dictionary<string, object>>
                {
                    new() { ["name"] = "Test1", ["active"] = "true" },
                    new() { ["name"] = "Test2", ["active"] = "false" },
                }
            });
            filesToValidate.Add((outPath, "[ACTIVE]"));
        }

        // Generate foreach_basic
        {
            var tpl = Tpl("@foreach {{items}}", "Item: {{name}}", "@endforeach");
            var outPath = Out("lo_foreach_basic.docx");
            MiniWord.SaveAsByTemplate(outPath, tpl, new Dictionary<string, object>
            {
                ["items"] = new List<Dictionary<string, object>>
                {
                    new() { ["name"] = "Alpha" },
                    new() { ["name"] = "Beta" },
                }
            });
            filesToValidate.Add((outPath, "Alpha"));
        }

        // Validate each file with LibreOffice
        foreach (var (file, expected) in filesToValidate)
        {
            var (success, pdfPath, error) = ValidateWithLibreOffice(file);
            _output.WriteLine($"LibreOffice {Path.GetFileName(file)}: {(success ? "OK" : "FAIL")} -> {expected}");
            if (!success)
                Assert.Fail($"LibreOffice could not open {file}: {error}");
            Assert.True(File.Exists(pdfPath), $"PDF not generated for {file}");
        }
    }

    #endregion

    #region Helpers

    private static string Tpl(params string[] lines)
    {
        var hash = string.Join("|", lines);
        var path = Path.Combine(OutputDir, $"tpl_{Math.Abs(hash.GetHashCode()):x}.docx");
        if (File.Exists(path)) return path;

        // Build document.xml body with paragraphs
        var paragraphs = lines.Select(l =>
            $"<w:p><w:pPr><w:pStyle w:val=\"PreformattedText\"/><w:bidi w:val=\"0\"/>" +
            $"<w:spacing w:before=\"0\" w:after=\"0\"/><w:jc w:val=\"left\"/></w:pPr>" +
            $"<w:r><w:rPr></w:rPr><w:t>{EscapeXml(l)}</w:t></w:r></w:p>").ToList();
        var bodyXml = string.Join("", paragraphs);

        // Use LibreOffice-created DOCX as skeleton (proper structure for OnlyOffice)
        var skeletonDocx = GetLibreOfficeSkeletonDocx();
        using var skeletonZip = new ZipArchive(new MemoryStream(skeletonDocx), ZipArchiveMode.Read);
        
        // Read all entries from skeleton
        var entries = new Dictionary<string, byte[]>();
        foreach (var entry in skeletonZip.Entries)
            entries[entry.FullName] = ReadEntry(entry);

        // Replace document.xml with our content
        var sectPr = "<w:sectPr><w:type w:val=\"nextPage\"/><w:pgSz w:w=\"11906\" w:h=\"16838\"/>" +
                     "<w:pgMar w:left=\"1134\" w:right=\"1134\" w:gutter=\"0\" w:header=\"0\" w:top=\"1134\" " +
                     "w:footer=\"0\" w:bottom=\"1134\"/></w:sectPr>";
        var docXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<w:document xmlns:o=\"urn:schemas-microsoft-com:office:office\" " +
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
            "xmlns:v=\"urn:schemas-microsoft-com:vml\" " +
            "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
            "xmlns:w10=\"urn:schemas-microsoft-com:office:word\" " +
            "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
            "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " +
            "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" " +
            "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" " +
            "xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" " +
            "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" " +
            "xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" " +
            "mc:Ignorable=\"w14 wp14 w15\">" +
            $"<w:body>{bodyXml}{sectPr}</w:body></w:document>";
        entries["word/document.xml"] = Encoding.UTF8.GetBytes(docXml);

        // Write the final DOCX
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Create);
        foreach (var kvp in entries)
        {
            var entry = zip.CreateEntry(kvp.Key, CompressionLevel.Optimal);
            using var entryStream = entry.Open();
            entryStream.Write(kvp.Value, 0, kvp.Value.Length);
        }
        return path;
    }

    // Cache for LibreOffice-created DOCX skeleton
    private static byte[]? _skeletonDocx;

    private static byte[] GetLibreOfficeSkeletonDocx()
    {
        if (_skeletonDocx != null) return _skeletonDocx;

        if (!DockerAvailable())
            throw new InvalidOperationException("Docker required for template creation");

        // Create a simple ODT file, convert to DOCX using LibreOffice in Docker
        var odtContent = "LINE1\nLINE2\nLINE3\n";
        var containerOdt = "/tmp/skeleton_template.odt";
        var containerDocx = "/tmp/skeleton_template.docx";

        // Write ODT content to container
        var psi = new ProcessStartInfo
        {
            FileName = "docker",
            Arguments = $"exec -i {DockerContainer} bash -c \"cat > {containerOdt}\"",
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using (var p = Process.Start(psi))
        {
            p!.StandardInput.Write(odtContent);
            p.StandardInput.Close();
            p.WaitForExit(5000);
        }

        // Convert ODT to DOCX using LibreOffice
        RunDocker($"exec {DockerContainer} libreoffice --headless --convert-to docx --outdir /tmp {containerOdt}");

        // Copy DOCX back to host
        var hostTemp = Path.Combine(OutputDir, "skeleton_template.docx");
        RunDocker($"cp {DockerContainer}:{containerDocx} {hostTemp}");

        // Clean up container
        RunDocker($"exec {DockerContainer} rm -f {containerOdt} {containerDocx}");

        _skeletonDocx = File.ReadAllBytes(hostTemp);
        return _skeletonDocx;
    }

    private static byte[] ReadEntry(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    private static string EscapeXml(string text) =>
        text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    private static string Out(string name) => Path.Combine(OutputDir, name);

    private static string Read(string path)
    {
        if (!File.Exists(path)) return "(file not found)";
        using var zip = new ZipArchive(File.OpenRead(path), ZipArchiveMode.Read);
        var entry = zip.GetEntry("word/document.xml");
        if (entry == null) return "(no document.xml)";
        using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
        var xml = reader.ReadToEnd();
        var texts = new List<string>();
        var idx = 0;
        while (true)
        {
            var start = xml.IndexOf("<w:t", idx, StringComparison.Ordinal);
            if (start < 0) break;
            var end = xml.IndexOf("</w:t>", start, StringComparison.Ordinal);
            if (end < 0) break;
            var content = xml.Substring(start, end - start);
            var gt = content.IndexOf('>');
            if (gt >= 0) texts.Add(content.Substring(gt + 1));
            idx = end + 6;
        }
        return string.Join("", texts);
    }

    // Validate a DOCX file by converting it to PDF using LibreOffice in Docker.
    // If LibreOffice can open and convert the file, it's valid for OnlyOffice too.
    private static (bool success, string pdfPath, string error) ValidateWithLibreOffice(string docxPath)
    {
        var fileName = Path.GetFileName(docxPath);
        var containerPath = $"/tmp/docx_validate/{fileName}";
        var containerDir = "/tmp/docx_validate";

        try
        {
            // Ensure container dir exists and is clean
            RunDocker($"exec {DockerContainer} bash -c \"mkdir -p {containerDir} && rm -f {containerDir}/*\"");

            // Copy DOCX to container
            RunDocker($"cp {docxPath} {DockerContainer}:{containerPath}");

            // Convert to PDF using LibreOffice headless
            var (exitCode, output, error) = RunDockerWithOutput(
                $"exec {DockerContainer} libreoffice --headless --convert-to pdf --outdir {containerDir} {containerPath}");

            if (exitCode != 0)
                return (false, "", $"LibreOffice exit {exitCode}: {error}");

            // Check PDF was created
            var pdfName = fileName.Replace(".docx", ".pdf");
            var containerPdfPath = $"{containerDir}/{pdfName}";
            // Use ls to check (more reliable than test -f across Docker boundary)
            var (lsExit, lsOutput, lsError) = RunDockerWithOutput(
                $"exec {DockerContainer} ls -1 {containerPdfPath}");
            if (lsExit != 0 || !lsOutput.Contains(pdfName))
                return (false, "", $"PDF not generated. Output: {output} | ls: {lsOutput} | err: {lsError}");

            // Copy PDF back to host for inspection
            var hostPdfPath = Path.Combine(OutputDir, pdfName);
            RunDocker($"cp {DockerContainer}:{containerPdfPath} {hostPdfPath}");

            return (true, hostPdfPath, "");
        }
        catch (Exception ex)
        {
            return (false, "", ex.Message);
        }
    }

    private static bool DockerAvailable()
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "docker",
                Arguments = $"ps --filter name={DockerContainer} --format {{{{.Names}}}}",
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            using var p = Process.Start(psi);
            p?.WaitForExit(5000);
            return p?.StandardOutput.ReadToEnd().Contains(DockerContainer) == true;
        }
        catch { return false; }
    }

    private static void RunDocker(string args)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "docker",
            Arguments = args,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var p = Process.Start(psi);
        p?.WaitForExit(30000);
    }

    private static (int exitCode, string output, string error) RunDockerWithOutput(string args)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "docker",
            Arguments = args,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        using var p = Process.Start(psi);
        p?.WaitForExit(30000);
        return (p?.ExitCode ?? -1, p?.StandardOutput.ReadToEnd() ?? "", p?.StandardError.ReadToEnd() ?? "");
    }

    #endregion
}