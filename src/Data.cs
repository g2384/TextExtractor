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
        public long Size { get; set; }
    }
}

