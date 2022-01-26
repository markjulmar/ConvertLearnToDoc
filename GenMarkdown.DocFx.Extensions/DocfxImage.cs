using System.IO;
using Julmar.GenMarkdown;

namespace GenMarkdown.DocFx.Extensions
{
    /// <summary>
    /// Generates a Docfx :::image::: extension
    /// </summary>
    public class DocfxImage : Image
    {
        /// <summary>
        /// Image type - defaults to "content".
        /// </summary>
        public string ImageType { get; set; }
        
        /// <summary>
        /// Localization scope
        /// </summary>
        public string LocScope { get; set; }

        /// <summary>
        /// True for a border
        /// </summary>
        public bool Border { get; set; }

        /// <summary>
        /// Lightbox image url
        /// </summary>
        public string Lightbox { get; set; }

        /// <summary>
        /// Optional link
        /// </summary>
        public string Link { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="altText"></param>
        /// <param name="imagePath"></param>
        /// <param name="description"></param>
        public DocfxImage(string altText, string imagePath, string description = "") : base(altText, imagePath, description)
        {
            ImageType = "content";
            LocScope = string.Empty;
            Lightbox = string.Empty;
            Link = string.Empty;
            Border = false;
        }

        /// <inheritdoc />
        public override void Write(TextWriter writer, MarkdownFormatting formatting)
        {
            writer.Write($":::image type=\"{ImageType}\"");
            writer.Write($" source=\"{ImagePath}\"");
            writer.Write($" alt-text=\"{AltText}\"");
            if (Border)
            {
                writer.Write($" border=\"{Border.ToString().ToLower()}\"");
            }

            if (!string.IsNullOrEmpty(Link))
            {
                writer.Write($" link=\"{Link.Trim()}\"");
            }
            
            if (!string.IsNullOrEmpty(Lightbox))
            {
                writer.Write($" lightbox=\"{Lightbox.Trim()}\"");
            }

            if (!string.IsNullOrEmpty(LocScope))
            {
                writer.Write($" loc-scope=\"{LocScope.ToLower()}\"");
            }

            if (!string.IsNullOrEmpty(Description))
            {
                writer.WriteLine(":::");
                writer.WriteLine(Description);
            }

            writer.WriteLine(":::");
            writer.WriteLine();
        }
    }
}