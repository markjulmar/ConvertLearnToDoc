using System.Collections.Generic;

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

        public bool HasContent => Lines.Count > 0;

        public UnitMetadata(string title)
        {
            Title = title;
        }
    }
}