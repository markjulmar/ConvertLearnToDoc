namespace ConvertLearnToDocWeb.Models
{
    public class LearnToDocModel
    {
        public string Organization { get; set; }
        public string Repository { get; set; }
        public string Branch { get; set; }
        public string Folder { get; set; }
        public string ZonePivot { get; set; }
        public bool EmbedNotebookData { get; set; }
    }
}
