using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LearnDocUtils
{
    internal class UnitMetadata
    {
        public string Title { get; }
        public List<string> Lines { get; } = new();
        public bool Sandbox { get; set; }
        public int LabId { get; set; }
        public string Notebook { get; set; }
        public string Interactivity { get; set; }

        public bool HasContent => Lines.Count > 0 && Lines.Any(s => !string.IsNullOrWhiteSpace(s));

        public UnitMetadata(string title)
        {
            Title = title;
        }

        public string BuildInteractivityOptions()
        {
            var sb = new StringBuilder();
            if (Sandbox)
                sb.AppendLine("sandbox: true");
            if (!string.IsNullOrEmpty(Interactivity))
                sb.AppendLine($"interactivity: {Interactivity}");
            if (LabId > 0)
                sb.AppendLine($"labId: {LabId}");
            if (!string.IsNullOrEmpty(Notebook))
                sb.AppendLine($"notebook: {Notebook}");
            
            return sb.ToString();
        }
    }
}