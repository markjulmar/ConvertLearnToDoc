using System;
using System.Linq;

namespace LearnDocUtils
{
    internal class ModuleMetadata
    {
        public string ModuleUid { get; set; }
        public string Title { get; set; }
        public string MsAuthor { get; set; }
        public string GitHubAlias { get; set; }
        public string Summary { get; set; }
        public DateTime LastModified { get; set; }
        public string MsTopic { get; set; }
        public string MsProduct { get; set; }
        public string Abstract { get; set; }
        public string Prerequisites { get; set; }
        public string IconUrl { get; set; }
        public string Levels { get; set; }
        public string Roles { get; set; }
        public string Products { get; set; }
        public string BadgeUid { get; set; }
        public string SEOTitle { get; set; }
        public string SEODescription { get; set; }

        public string GetList(string cdv, string defaultValue)
        {
            if (string.IsNullOrEmpty(cdv))
                return defaultValue;
            string[] levels = cdv.Split(',');
            return string.Join(Environment.NewLine, levels.Select(s => "- " + s));
        }
    }
}
