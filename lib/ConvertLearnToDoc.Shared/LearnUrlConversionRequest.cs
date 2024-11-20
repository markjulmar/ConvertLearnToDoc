using System.ComponentModel.DataAnnotations;

namespace ConvertLearnToDoc.Shared;

public sealed class LearnUrlConversionRequest
{
    [Required, MaxLength(200)]
    public string Url { get; set; } = string.Empty;
    public string? ZonePivot { get; set; }
    public bool EmbedNotebooks { get; set; }

    public bool IsValid() => !string.IsNullOrWhiteSpace(Url);

    public override string ToString()
    {
        return $"{Url}, ZonePivot={ZonePivot}, EmbedNotebooks={EmbedNotebooks}";
    }
}