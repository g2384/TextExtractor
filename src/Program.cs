using Serilog;
using Serilog.Sinks.SystemConsole.Themes;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
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
            File.WriteAllText(outputPath, $"");
            var dataList = new List<Data>();
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

