using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace YamlMimeModule;

public class Root
{
    public string? Uid { get; set; }
    public string? Title { get; set; }
    public string? Summary { get; set; }
    public string? Abstract { get; set; }
    public string? Prerequisites { get; set; }
    public string? IconUrl { get; set; }
    public List<string>? Levels { get; set; }
    public List<string>? Roles { get; set; }
    public List<string>? Products { get; set; }
    public List<string>? Units { get; set; }
    
    public Badge? Badge { get; set; }
    public Metadata? Metadata { get; set; }

}
public class Badge
{
    public string Uid { get; set; }
}

public class Metadata
{
    public string? Title { get; set; }

    public string? Description { get; set; }

    [YamlMember(Alias = "ms.date", ApplyNamingConventions = false)]
    public string? MsDate { get; set; }

    public string? Author { get; set; }

    [YamlMember(Alias = "ms.author", ApplyNamingConventions = false)]
    public string? MsAuthor { get; set; }

    [YamlMember(Alias = "ms.topic", ApplyNamingConventions = false)]
    public string? MsTopic { get; set; }

    [YamlMember(Alias = "ms.prod", ApplyNamingConventions = false)]
    public string? MsProd { get; set; }

    [YamlMember(Alias = "ROBOTS", ApplyNamingConventions = false)]
    public string? Robots { get; set; }
}