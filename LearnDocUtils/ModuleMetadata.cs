using System;

namespace ConvertDocToLearn
{
    internal class ModuleMetadata
    {
        public string ModuleUid { get; set; }
        public string Title { get; set; }
        public string MsAuthor { get; set; }
        public string Summary { get; set; }
        public DateTime LastModified { get; set; }
        public string MsTopic { get; set; }
        public string MsProduct { get; set; }
        public string Abstract { get; set; }
    }
}
