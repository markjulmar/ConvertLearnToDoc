using System;
using System.Collections.Generic;
using System.Linq;
using MSLearnRepos;

namespace LearnDocUtils
{
    public class ModuleMetadata
    {
        private readonly TripleCrownModule moduleData;
        public TripleCrownModule ModuleData => moduleData;
        public ModuleMetadata(TripleCrownModule moduleData)
        {
            this.moduleData = moduleData ?? new TripleCrownModule();
            this.moduleData.Metadata ??= new TripleCrownMetadata();
        }

        public static string GetList(List<string> items, string defaultValue) =>
            items == null || items.Count == 0
                ? defaultValue
                : string.Join(Environment.NewLine, items.Select(s => "- " + s));
    }
}
