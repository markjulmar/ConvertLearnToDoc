using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using Markdig.Extensions.Yaml;
using Markdig.Renderer.Docx.Blocks;
using Markdig.Renderer.Docx.Inlines;
using CodeBlockRenderer = Markdig.Renderer.Docx.Blocks.CodeBlockRenderer;
using HeadingRenderer = Markdig.Renderer.Docx.Blocks.HeadingRenderer;
using ListRenderer = Markdig.Renderer.Docx.Blocks.ListRenderer;
using ParagraphRenderer = Markdig.Renderer.Docx.Blocks.ParagraphRenderer;
using QuoteBlockRenderer = Markdig.Renderer.Docx.Blocks.QuoteBlockRenderer;
using TripleColonInlineRenderer = Markdig.Renderer.Docx.Inlines.TripleColonInlineRenderer;

namespace Markdig.Renderer.Docx;

/// <summary>
/// Options for the renderer
/// </summary>
public class DocxRendererOptions
{
    /// <summary>
    /// Zone pivot to render (null/empty for all)
    /// </summary>
    public string ZonePivot { get; set; }

    /// <summary>
    /// Function to retrieve a file referenced by an element.
    /// </summary>
    public Func<MarkdownObject,string,byte[]> ReadFile { get; set; }

    /// <summary>
    /// Convert a relative URL found in content to an absolute URL for the document.
    /// </summary>
    public Func<string, string> ConvertRelativeUrl { get; set; }

    /// <summary>
    /// Optional logger function
    /// </summary>
    public Action<string> Logger { get; set; }

}

/// <summary>
/// DoxC renderer for a Markdown <see cref="MarkdownDocument"/> object.
/// </summary>
public class DocxObjectRenderer : IDocxRenderer
{
    private readonly IDocument document;
    private readonly List<IDocxObjectRenderer> renderers;
    private readonly string moduleFolder;
    private readonly DocxRendererOptions options;
    private MarkdownDocument markdownDocument;

    /// <summary>
    /// This holds elements where previous inline renderers had to reach into the stream
    /// and render siblings. It's used to avoid double rendering.
    /// </summary>
    public IList<MarkdownObject> OutOfPlaceRendered { get; }

    public string ZonePivot => options?.ZonePivot;

    public DocxObjectRenderer(IDocument document, string moduleFolder, DocxRendererOptions options)
    {
        this.moduleFolder = moduleFolder;
        this.options = options;
        this.document = document;
        this.OutOfPlaceRendered = new List<MarkdownObject>();

        renderers = new List<IDocxObjectRenderer>
        {
            // Ignored blocks
            new IgnoredBlock(typeof(YamlFrontMatterBlock)),

            // Block handlers
            new HeadingRenderer(),
            new ParagraphRenderer(),
            new ListRenderer(),
            new QuoteBlockRenderer(),
            new QuoteSectionNoteRenderer(),
            new CodeBlockRenderer(),
            new TripleColonRenderer(),
            new TableRenderer(),
            new InclusionRenderer(),
            new LinkReferenceDefinitionGroupRenderer(),
            new HtmlBlockRenderer(),
            new MonikerRangeRenderer(),
            new ThematicBreakRenderer(),
            new RowBlockRenderer(),

            // Inline handlers
            new LiteralInlineRenderer(),
            new EmphasisInlineRenderer(),
            new LineBreakInlineRenderer(),
            new LinkInlineRenderer(),
            new AutolinkInlineRenderer(),
            new CodeInlineRenderer(),
            new DelimiterInlineRenderer(),
            new HtmlEntityInlineRenderer(),
            new LinkReferenceDefinitionRenderer(),
            new TaskListRenderer(),
            new HtmlInlineRenderer(),
            new NolocInlineRenderer(),
            new TripleColonInlineRenderer()
        };
    }

    public IDocxObjectRenderer FindRenderer(MarkdownObject obj)
    {
        var renderer = renderers.FirstOrDefault(r => r.CanRender(obj));
        if (renderer == null && options?.Logger != null)
        {
            var type = obj.GetType();
            var sb = new StringBuilder($"Missing renderer for {type}:").AppendLine();
            foreach (var pi in type.GetProperties())
            {
                try
                {
                    sb.AppendLine($"\t{pi.Name}=\"{pi.GetValue(obj)}\"");
                }
                catch
                {
                    // Ignore
                }
            }

            options?.Logger?.Invoke(sb.ToString());
        }

        return renderer;
    }

    public void Render(MarkdownDocument mdDoc)
    {
        this.markdownDocument = mdDoc ?? throw new ArgumentNullException(nameof(mdDoc));

        for (var index = 0; index < markdownDocument.Count; index++)
        {
            var block = markdownDocument[index];

            // Special case RowBlock and children to generate a full table.
            // This is an optimization when the RowBlock is a root element in the document.
            // If it's contained in some other block (like a List), then the default handler will kick in.
            if (block is RowBlock)
            {
                var rows = new List<RowBlock>();
                do
                {
                    rows.Add((RowBlock) block);
                    block = index+1 < markdownDocument.Count ? markdownDocument[++index] : null;
                } while (block is RowBlock);

                new RowBlockRenderer().Write(this, document, rows);
                continue;
            }

            // Find the renderer and process.
            var renderer = FindRenderer(block);

            try
            {
                renderer?.Write(this, document, null, block);
            }
            catch (AggregateException aex)
            {
                var ex = aex.Flatten();
                options?.Logger?.Invoke($"{ex.GetType().Name}: {ex.Message}");
                options?.Logger?.Invoke(ex.StackTrace);
            }
            catch (Exception ex)
            {
                options?.Logger?.Invoke($"{ex.GetType().Name}: {ex.Message}");
                options?.Logger?.Invoke(ex.StackTrace);
            }
        }
    }

    public string ConvertRelativeUrl(string url) => options?.ConvertRelativeUrl(url);

    public byte[] GetFile(MarkdownObject source, string path) => options?.ReadFile?.Invoke(source, path);

    public void AddComment(Paragraph owner, string commentText)
    {
        string user = Environment.UserInteractive ? Environment.UserName : Environment.GetEnvironmentVariable("CommentUserName");
        if (string.IsNullOrEmpty(user))
            user = "Office User";

        owner.AttachComment(document.CreateComment(user, commentText));
    }

    public Drawing InsertImage(Paragraph currentParagraph, MarkdownObject owner, string imageUrl, string altText, string title, bool hasBorder)
    {
        string path = ResolvePath(moduleFolder, imageUrl);
        DXPlus.Image image = null;

        // Simple fix for images in the same folder as the Markdown content - in this case we need to be able
        // to distinguish them from images added later with no path. We'll add a path here which we can catch
        // later to know these were part of the original content graph.
        if (Path.GetDirectoryName(imageUrl) == string.Empty)
        {
            Debug.Assert(!imageUrl.Contains('/') && !imageUrl.Contains('\\'));
            imageUrl = "./" + imageUrl;
        }

        if (IsInternetUrl(path))
        {
            // Copy to a local stream -- the AddImage does some position management
            // which fails on the HTTP stream.
            using var ms = new MemoryStream();
            using (var stream = new HttpClient().GetStreamAsync(imageUrl).Result)
            {
                stream.CopyTo(ms);
            }

            image = document.CreateImage(ms, DetermineContentTypeFromUrl(imageUrl));
        }

        else if (!File.Exists(path))
        {
            byte[] contents = GetFile(owner, imageUrl);
            if (contents != null)
            {
                image = document.CreateImage(new MemoryStream(contents, false), DetermineContentTypeFromUrl(imageUrl));
            }
        }
        else
        {
            image = document.CreateImage(path);
        }

        if (image != null)
        {
            var picture = image.CreatePicture(imageUrl, altText);
            if (picture.Width > 600 && picture.Height != null)
            {
                double ratio = picture.Height.Value / picture.Width.Value;
                picture.Drawing.Width = 600; picture.Width = 600;
                var height = Math.Round(600 * ratio);
                picture.Drawing.Height = height; picture.Height = height;
            }

            if (hasBorder)
                picture.BorderColor = Color.DarkGray;

            picture.Name = Path.GetFileName(imageUrl);
            picture.Description = altText;
            currentParagraph.Add(picture);

            if (!string.IsNullOrEmpty(title))
                picture.Drawing.AddCaption(": " + title);

            return picture.Drawing;
        }
            
        options?.Logger?.Invoke($"Error: unable to add image {imageUrl} to document.");
        return null;
    }

    private string DetermineContentTypeFromUrl(string imageUrl)
    {
        if (IsInternetUrl(imageUrl))
        {
            var uri = new Uri(imageUrl, UriKind.Absolute);
            imageUrl = uri.GetLeftPart(UriPartial.Path);
        }

        imageUrl = Path.GetExtension(imageUrl.ToLower());

        string contentType = imageUrl switch
        {
            ".apng" => "image/apng",
            ".avif" => "image/avif",
            ".bmp" => "image/bmp",
            ".ico" => "image/x-icon",
            ".cur" => "image/x-icon",
            ".png" => "image/png",
            ".jpg" => "image/jpeg",
            ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".svg" => "image/svg+xml",
            ".tif" => "image/tiff",
            ".tiff" => "image/tiff",
            _ => null
        };

        if (contentType == null)
        {
            options?.Logger?.Invoke($"Error: unable to determine content type for {imageUrl}.");
        }

        return contentType;
    }

    private static bool IsInternetUrl(string path) => path?.ToLower().StartsWith("http") == true;

    /// <summary>
    /// Returns a specific embedded resource by name.
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    public Stream GetEmbeddedResource(string name) => Assembly.GetExecutingAssembly().GetManifestResourceStream("Markdig.Renderer.Docx.Resources."+name);

    /// <summary>
    /// Resolve the image path to our local root folder.
    /// </summary>
    /// <param name="rootFolder">Folder we downloaded images to</param>
    /// <param name="path">Image path in content</param>
    /// <returns>Resolved path</returns>
    private static string ResolvePath(string rootFolder, string path)
    {
        if (rootFolder == null) throw new ArgumentNullException(nameof(rootFolder));
        if (path == null) throw new ArgumentNullException(nameof(path));

        path = path.Trim('\"');

        if (path.ToLower().StartsWith(@"..\media")
            || path.ToLower().StartsWith("../media"))
        {
            path = path[1..];
        }

        if (IsInternetUrl(path)) return path;

        if (!Path.IsPathRooted(path))
        {
            path = Path.Combine(rootFolder, path);
        }

        int index = path.IndexOf('#');
        if (index>0)
            path = path[..index];

        return path;
    }
}