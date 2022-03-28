using Serilog;
using Serilog.Sinks.SystemConsole.Themes;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using Toxy;

namespace TextExtractor
{
    public class Data
    {
        public string FileName { get; set; }
        public string Extension { get; set; }
        public string FullPath { get; set; }
        public string Directory { get; set; }
        public string ModifiedTime { get; set; }
        public string[] Paragraphs { get; set; }
    }
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
            var outputPath = DateTime.Now.ToString("yyyyMMddhhmmss") + "_output.html";
            var html = File.ReadAllText("output.html");
            File.WriteAllText(outputPath, $"");
            var dataList = new List<Data>();
            foreach (var file in allFiles)
            {
                var fi = new FileInfo(file);
                var ext = fi.Extension;
                var lastWriteTIme = fi.LastWriteTime;
                Log.Information($"Processing {fi.Name}");
                switch (ext)
                {
                    case ".docx":
                    case ".doc":
                    case ".pdf":
                        var context = new ParserContext(file);
                        var parser = ParserFactory.CreateDocument(context);
                        var result = parser.Parse();
                        var paragraphs = result.Paragraphs.Select(e => e.Text.Trim()).Where(e => !string.IsNullOrWhiteSpace(e)).ToArray();
                        var d = new Data
                        {
                            Extension = ext,
                            FileName = fi.Name,
                            FullPath = fi.FullName,
                            Paragraphs = paragraphs,
                            Directory = fi.Directory.FullName,
                            ModifiedTime = lastWriteTIme.ToString("yyy-MM-dd HH:mm:ss")
                        };
                        dataList.Add(d);
                        break;
                    default:
                        Log.Error($"{ext} is not supported.");
                        break;
                }
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
    }
}

