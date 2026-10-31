using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System.Linq;

foreach (var path in args)
{
    using var doc = PresentationDocument.Open(path, false);
    var v = new OpenXmlValidator(FileFormatVersions.Microsoft365);
    var errors = v.Validate(doc).ToList();
    System.Console.WriteLine($"{System.IO.Path.GetFileName(path)}: {errors.Count} schema errors");
    foreach (var g in errors.GroupBy(e => e.Description).OrderByDescending(g => g.Count()).Take(12))
    {
        var e = g.First();
        var desc = g.Key.Length > 170 ? g.Key.Substring(0, 170) : g.Key;
        System.Console.WriteLine($"  x{g.Count()} [{e.Part?.Uri}] {desc}");
        System.Console.WriteLine($"      at {e.Path?.XPath?.Split('/').LastOrDefault()}");
    }
}
