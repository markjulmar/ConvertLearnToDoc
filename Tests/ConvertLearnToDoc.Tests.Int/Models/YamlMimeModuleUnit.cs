using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace YamlMimeModuleUnit;

public class Root
{
    public string? Uid { get; set; }
    public string? Title { get; set; }
    public Metadata? Metadata { get; set; }
    public int? DurationInMinutes { get; set; }
    public bool? Sandbox { get; set; }
    public bool? AzureSandbox { get; set; }
    public string? Content { get; set; }
    public string? Interactive { get; set; }
    public string? Interactivity { get; set; }
    public Quiz? Quiz { get; set; }

    
}


public class Quiz
{
    public string? Title { get; set; }
    public List<Question>? Questions { get; set; }
}

public class Question
{
    public string? Content { get; set; }
    public List<Choice>? Choices { get; set; }
}
public class Choice
{
    public string? Content { get; set; }
    public bool? IsCorrect { get; set; }
    public string? Explanation { get; set; }
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

    public override bool Equals(object? obj)
    {
        return obj is Metadata metadata &&
               Title == metadata.Title &&
               Description == metadata.Description &&
               MsDate == metadata.MsDate &&
               Author == metadata.Author &&
               MsAuthor == metadata.MsAuthor &&
               MsTopic == metadata.MsTopic &&
               MsProd == metadata.MsProd &&
               Robots == metadata.Robots;
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Title, Description, MsDate, Author, MsAuthor, MsTopic, MsProd, Robots);
    }
}

