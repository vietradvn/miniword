using MiniSoftware;
using System.IO.Compression;
using System.Text;

var templatePath = "/home/ta/Downloads/template.docx";
var outputPath = "/home/ta/Downloads/template_filled.docx";

var vars = new Dictionary<string, object>
{
    ["patient_name"] = "Nguyễn Văn An",
    ["dob"] = "15/03/1985",
    ["gender"] = "Male",
    ["patient_id"] = "PT-2026-001234",
    ["accession_number"] = "ACC-78901",
    ["order_name"] = "CT Head without contrast",
    ["exam_date"] = "2026-06-18",
    ["referring_physician"] = "Dr. Trần Quang Minh",
    ["conclusion"] = "No acute intracranial abnormality detected.",
    ["radiologist_name"] = "Dr. Lê Thị Hương",
    ["is_vip"] = "true",
    ["has_addendum"] = "true",
    ["addendum_content"] = "Follow-up recommended in 6 months.",
    ["results"] = new List<Dictionary<string, object>>
    {
        new() { ["stt"] = 1, ["name"] = "CT Head", ["result"] = "Normal" },
        new() { ["stt"] = 2, ["name"] = "CT Neck", ["result"] = "Mild degenerative changes" },
        new() { ["stt"] = 3, ["name"] = "CT Sinus", ["result"] = "Clear" },
    },
    ["findings"] = new List<Dictionary<string, object>>
    {
        new() { ["name"] = "Brain parenchyma", ["value"] = "Normal density, no mass effect" },
        new() { ["name"] = "Ventricular system", ["value"] = "Normal size and position" },
        new() { ["name"] = "Basal cisterns", ["value"] = "Patent" },
        new() { ["name"] = "Calvarium", ["value"] = "Intact" },
    },
};

MiniWord.SaveAsByTemplate(outputPath, templatePath, vars);
Console.WriteLine($"Output: {outputPath}");
