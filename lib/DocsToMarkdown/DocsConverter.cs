using System.Diagnostics;
using System.Text;

namespace Julmar.DocsToMarkdown;

public enum FileType
{
    Markdown,
    Yaml,
    Image,
    Asset,
    Folder,
}

public record FileEntry(string Filename, FileType FileType);

public class DocsConverter(string rootFolder, Uri url)
{
    private List<FileEntry> CreatedFiles { get; } = [];
    private readonly List<Task> _downloadTasks = [];
    private string? _moduleName;
    private string _rootFolder = rootFolder;

    public async Task<IReadOnlyCollection<FileEntry>> ConvertAsync(bool useExtensions = true, Action<string>? logger = null)
    {
        var converter = await IHtmlConverter.LoadAsync(url,
            new HtmlConverterOptions()
            {
                UseDocfxExtensions = useExtensions,
                UnhandledTagCallback = logger ?? (t => Debug.WriteLine($"Unhandled tag: {t}")),
                DownloadAsset = (href,assetUrl) => _downloadTasks.Add(OnDownloadAsset(href, assetUrl))
            });

        if (converter.IsTrainingModulePage)
            await ProcessTrainingModule(converter);
        else
        {
            var filename = Path.Combine(_rootFolder,
                Path.ChangeExtension(converter.Url.Segments.Last().TrimEnd('/'), ".md"));
            await ProcessSinglePage(converter, filename);
        }

        // Wait for any downloads
        await Task.WhenAll(_downloadTasks.ToArray());

        return CreatedFiles;
    }

    private async Task ProcessTrainingModule(IHtmlConverter converter)
    {
        _moduleName = converter.Url.Segments.Last().TrimEnd('/');
        var folder = Path.Combine(_rootFolder, _moduleName);
        if (!Directory.Exists(folder))
        {
            Directory.CreateDirectory(folder);
            CreatedFiles.Add(new FileEntry(folder, FileType.Folder));
        }

        _rootFolder = folder;

        // Write the index file.
        var indexFile = Path.Combine(folder, "index.yml");
        await File.WriteAllTextAsync(indexFile, converter.Metadata.YamlHeader);
        CreatedFiles.Add(new FileEntry(indexFile, FileType.Yaml));

        var includesFolder = Path.Combine(folder, "includes");
        if (!Directory.Exists(includesFolder))
        {
            Directory.CreateDirectory(includesFolder);
            CreatedFiles.Add(new FileEntry(includesFolder, FileType.Folder));
        }

        // Write each unit
        await foreach(var unit in converter.GetTrainingUnitsConverter())
        {
            // Write the YAML
            var filename = Path.ChangeExtension(unit.Url.Segments.Last().TrimEnd('/'), "yml");
            var yamlFile = Path.Combine(folder, filename);

            var mdFilename = Path.ChangeExtension(filename, "md");
            var sb = new StringBuilder(unit.Metadata.YamlHeader);
            sb.AppendLine("content: |")
                .AppendLine($"  [!include[](includes/{mdFilename})]");
            
            await File.WriteAllTextAsync(yamlFile, sb.ToString());
            CreatedFiles.Add(new FileEntry(yamlFile, FileType.Yaml));

            // Write the Markdown
            var mdFile = Path.Combine(includesFolder, mdFilename);
            await ProcessSinglePage(unit, mdFile, writeYaml: false);
        }
    }

    private async Task ProcessSinglePage(IHtmlConverter converter, string filename, bool writeYaml = true)
    {
        await using var writer = new StreamWriter(filename);

        if (writeYaml)
        {
            // Write the YAML header
            var header = converter.Metadata.YamlHeader;
            if (!string.IsNullOrEmpty(header))
            {
                var yamlDivider = new string('-', 3);
                await writer.WriteLineAsync(yamlDivider);
                await writer.WriteLineAsync(header);
                await writer.WriteLineAsync(yamlDivider);
                await writer.WriteLineAsync();
            }
        }
        
        // Write the Markdown content
        foreach (var line in converter.GetMarkdown())
            await writer.WriteLineAsync(line);
        
        CreatedFiles.Add(new FileEntry(filename, FileType.Markdown));
    }

    private static readonly string[] ImageExtensions = [".jpg", ".jpeg", ".png", ".gif", ".svg", ".bmp", "ico"];
    private async Task OnDownloadAsset(string source, Uri assetUrl)
    {
        // Ignore if it's a full URL.
        if (source.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            return;

        // Do a filename check BEFORE the async download.
        var filename = Path.Combine(_rootFolder, source);
        if (File.Exists(filename)
            || CreatedFiles.Any(f => f.Filename == filename))
            return;

        var folder = Path.GetDirectoryName(filename);
        if (!string.IsNullOrEmpty(folder) && !Directory.Exists(folder))
        {
            Directory.CreateDirectory(folder);
            CreatedFiles.Add(new FileEntry(folder, FileType.Folder));
        }

        var isImage = Array.IndexOf(ImageExtensions, Path.GetExtension(filename).ToLower()) >= 0;
        CreatedFiles.Add(new FileEntry(filename, isImage ? FileType.Image : FileType.Asset));

        try
        {
            using var client = new HttpClient();
            await using var stm = await client.GetStreamAsync(assetUrl);
            await using var file = File.Create(filename);
            await stm.CopyToAsync(file);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error downloading {assetUrl}: {ex.Message}");
        }
    }
}