using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Security.Cryptography;
using System.Text.Json;

if (args.Length != 1 || !Directory.Exists(args[0])) throw new ArgumentException("Pass the directory containing synthetic XLSX fixtures.");
var files = Directory.GetFiles(args[0], "*.xlsx").OrderBy(x => x).ToArray();
if (files.Length == 0) throw new InvalidOperationException("No XLSX fixtures were generated.");
var validator = new OpenXmlValidator(FileFormatVersions.Office2021) { MaxNumberOfErrors = 500 };
var reports = new List<object>();
var failures = 0;
foreach (var file in files)
{
    var before = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(file)));
    try
    {
        using (var document = SpreadsheetDocument.Open(file, false, new OpenSettings { AutoSave = false }))
        {
            var errors = validator.Validate(document).Select(e => new { e.Id, type = e.ErrorType.ToString(), e.Description, part = e.Part?.Uri.ToString(), path = e.Path?.XPath }).ToArray();
            if (errors.Length > 0) failures++;
            reports.Add(new { file = Path.GetFileName(file), errors, unchanged = before == Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(file))) });
            Console.WriteLine($"{Path.GetFileName(file)}: {errors.Length} schema/semantic errors");
        }
        if (before != Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(file)))) throw new InvalidOperationException("Read-only validation changed a fixture.");
    }
    catch (Exception error)
    {
        failures++;
        reports.Add(new { file = Path.GetFileName(file), exception = error.ToString() });
    }
}
File.WriteAllText(Path.Combine(args[0], "sdk-validation.json"), JsonSerializer.Serialize(new { sdk = "DocumentFormat.OpenXml 3.5.1", format = "Office2021", scope = "Read-only OOXML schema and SDK semantic validation, not Microsoft Excel execution or rendering", files = files.Length, failures, reports }, new JsonSerializerOptions { WriteIndented = true }));
return failures == 0 ? 0 : 1;
