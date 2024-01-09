using MSLearnRepos;
using Newtonsoft.Json.Serialization;
using Newtonsoft.Json;
using Document = DXPlus.Document;
using ConvertLearnToDoc.Shared;

namespace LearnDocUtils
{
    public static class DocsDownloader
    {
        public static async Task<string> DownloadFileAsync(ILearnRepoService learnRepo, string inputFile, string outputFolder)
        {
            var (_, text) = await learnRepo.ReadFileForPathAsync(inputFile);

            string filename = Path.GetFileName(inputFile);
            string outputFile = Path.Combine(outputFolder, filename);

            await File.WriteAllTextAsync(outputFile, text);
            return outputFile;
        }

        public static Dictionary<object,object> GetArticleMetadataFromDocument(string docxFile)
        {
            using var doc = Document.Load(docxFile);

            Dictionary<object, object> metadata = null;

            if (doc.CustomProperties.TryGetValue(nameof(Module.Metadata), out var property) && property != null)
            {
                var text = property.Value;
                if (text?.Length > 0)
                {
                    // Try a Learn module first. We will cull values we don't want/need.
                    try
                    {
                        var moduleData = PersistenceUtilities
                            .JsonStringToObject<Module>(text);
                        if (moduleData is { Uid: not null })
                        {
                            metadata = new();
                            if (!string.IsNullOrWhiteSpace(moduleData.Uid))
                                metadata["uid"] = moduleData.Uid;
                            if (!string.IsNullOrWhiteSpace(moduleData.Title))
                                metadata["title"] = moduleData.Title;
                            if (!string.IsNullOrWhiteSpace(moduleData.Summary))
                                metadata["description"] = moduleData.Summary;
                            if (!string.IsNullOrWhiteSpace(moduleData.Metadata?.Author))
                                metadata["author"] = moduleData.Metadata?.Author;
                            if (!string.IsNullOrWhiteSpace(moduleData.Metadata?.MsAuthor))
                                metadata["ms.author"] = moduleData.Metadata?.MsAuthor;
                            if (!string.IsNullOrWhiteSpace(moduleData.Metadata?.MsDate))
                                metadata["ms.date"] = moduleData.Metadata?.MsDate;
                        }
                    }
                    catch
                    {
                        // Ignore
                    }

                    if (metadata == null)
                    {
                        try
                        {
                            metadata = PersistenceUtilities.JsonStringToDictionary(text);
                        }
                        catch
                        {
                            // Ignore
                        }
                    }
                }
            }

            metadata ??= new Dictionary<object, object>();

            if (!string.IsNullOrWhiteSpace(doc.Properties.Title))
                metadata["title"] = doc.Properties.Title;
            if (!string.IsNullOrWhiteSpace(doc.Properties.Subject))
                metadata["description"] = doc.Properties.Subject;
            if (!string.IsNullOrWhiteSpace(doc.Properties.Creator))
                metadata["author"] = doc.Properties.Creator;
            if (!string.IsNullOrWhiteSpace(doc.Properties.LastSavedBy))
                metadata["author"] = doc.Properties.LastSavedBy;

            // Use SaveDate first, then CreatedDate if unavailable.
            if (doc.Properties.SaveDate != null)
                metadata["ms.date"] = doc.Properties.SaveDate.Value.ToString("MM/dd/yyyy");
            else if (doc.Properties.CreatedDate != null)
                metadata["ms.date"] = doc.Properties.CreatedDate.Value.ToString("MM/dd/yyyy");
            else
                metadata["ms.date"] = DateTime.Now.ToString("MM/dd/yyyy");

            return metadata;
        }
    }
}
