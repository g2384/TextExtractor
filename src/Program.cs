using System.Text;
using Serilog;
using Serilog.Sinks.SystemConsole.Themes;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Toxy;

namespace TextExtractor
{
    public static class Program
    {
        public static void Main(params string[] paths)
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console(outputTemplate: "[{Timestamp:yyy-MM-dd HH:mm:ss} {Level:w4}] {Message:lj}{NewLine}{Exception}", theme: AnsiConsoleTheme.Code)
                .CreateLogger();

            if (paths?.Length == 0)
            {
                Log.Error("Provide a path.");
                return;
            }
            var path = paths[0];
            var allFiles = FileHelper.GetAllFiles(path, "*.*").ToArray();
            var outputPath = DateTime.Now.ToString("yyyyMMddHHmmss") + "_output.html";
            var html = File.ReadAllText("output.html");
            var dataList = new List<Data>();
            var registeredCoding = false;
            foreach (var file in allFiles)
            {
                var fi = new FileInfo(file);
                var ext = fi.Extension.ToLowerInvariant();
                var lastWriteTIme = fi.LastWriteTime;
                Log.Information($"Processing {fi.Name}");
                var paragraphs = Array.Empty<string>();
                switch (ext)
                {
                    case ".docx":
                    case ".doc":
                    case ".pdf":
                        paragraphs = GetParagraphs(file);
                        break;
                    case ".xlsx":
                    case ".xls":
                        paragraphs = GetParagraphsExcel(file);
                        break;
                    case ".pptx":
                    case ".ppt":
                        paragraphs = GetParagraphsSlides(file);
                        break;
                    case ".msg":
                        if (!registeredCoding)
                        {
                            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                            registeredCoding = true;
                        }
                        paragraphs = GetParagraphsMsg(file);
                        break;
                    default:
                        Log.Error($"{ext} is not supported.");
                        break;
                }
                var d = new Data
                {
                    Extension = ext,
                    FileName = Path.GetFileNameWithoutExtension(fi.Name),
                    FullPath = fi.FullName,
                    Paragraphs = paragraphs,
                    Directory = fi.Directory.FullName,
                    ModifiedTime = lastWriteTIme.ToString("yyy-MM-dd HH:mm:ss")
                };
                dataList.Add(d);
            }
            var json = JsonSerializer.Serialize(dataList, new JsonSerializerOptions
            {
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
            html = html.Replace("$data$", json);
            File.WriteAllText(outputPath, html);
        }

        private static Regex lineFeedRegex = new Regex("[\r\n]+", RegexOptions.Compiled);

        private static string[] GetParagraphsMsg(string file)
        {
            var context = new ParserContext(file);
            var parser = ParserFactory.CreateEmail(context);
            var result = parser.Parse();
            var l1 = nameof(result.ArrivalTime) + ": " + result.ArrivalTime?.ToString("yyy-MM-dd HH:mm:ss");
            var l2 = nameof(result.Attachments) + ": " + string.Join(", ", result.Attachments);
            var l3 = nameof(result.Bcc) + ": " + string.Join(", ", result.Bcc);
            var l4 = nameof(result.Cc) + ": " + string.Join(", ", result.Cc);
            var l5 = nameof(result.From) + ": " + result.From;
            var l6 = nameof(result.To) + ": " + string.Join(", ", result.To);
            var l7 = nameof(result.Subject) + ": " + result.Subject;
            var text = lineFeedRegex.Split(result.TextBody);
            text = text.Select(e => e.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).ToArray();
            return new[] { l1, l2, l3, l4, l5, l6, l7 }.Concat(text).ToArray();
        }

        private static string[] GetParagraphsSlides(string file)
        {
            var context = new ParserContext(file);
            var parser = ParserFactory.CreateSlideshow(context);
            var result = parser.Parse();
            var text = result.Slides.SelectMany(e => e.Texts).Select(e => e.Trim());
            var paragraphs = text.Where(e => !string.IsNullOrWhiteSpace(e)).ToArray();
            return paragraphs;
        }

        private static string[] GetParagraphsExcel(string file)
        {
            var context = new ParserContext(file);
            var parser = ParserFactory.CreateSpreadsheet(context);
            var result = parser.Parse();
            var cells = result.Tables.SelectMany(e => e.Rows).SelectMany(e => e.Cells);
            var paragraphs = cells.Select(e =>
            {
                if (string.IsNullOrWhiteSpace(e.Formula))
                {
                    return e.Value.Trim();
                }
                return e.Formula.Trim() + "(" + e.Value.Trim() + ")";
            }).Where(e => !string.IsNullOrWhiteSpace(e)).ToArray();
            return paragraphs;
        }

        private static string[] GetParagraphs(string file)
        {
            var context = new ParserContext(file);
            var parser = ParserFactory.CreateDocument(context);
            var result = parser.Parse();
            var paragraphs = result.Paragraphs.Select(e => e.Text.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).ToArray();
            return paragraphs;
        }
    }
}

