﻿using Microsoft.AspNetCore.Components.Forms;

namespace ConvertLearnToDoc.Shared
{
    public class BrowserFile
    {
        public string ContentType { get; set; } = null!;
        public DateTimeOffset LastModified { get; set; }
        public string FileName { get; set; } = null!;
        public byte[] Contents { get; set; } = null!;

        public static async Task<BrowserFile?> CreateAsync(IBrowserFile selectedFile)
        {
            var bf = new BrowserFile
            {
                ContentType = selectedFile.ContentType,
                LastModified = selectedFile.LastModified,
                FileName = selectedFile.Name
            };

            await using var stream = selectedFile.OpenReadStream();
            MemoryStream ms = new();
            await stream.CopyToAsync(ms);
            bf.Contents = ms.ToArray();

            return bf;
        }

        public override string ToString()
        {
            return $"FileName={FileName}, Size={Contents.Length}, LastModified={LastModified}, Type={ContentType}";
        }
    }

    public class ArticleOrModuleRef
    {
        public bool IsArticle { get; set; }
        public BrowserFile? Document { get; set; }
        public bool UsePlainMarkdown { get; set; }
        public bool UseAsterisksForBullets { get; set; }
        public bool UseAsterisksForEmphasis { get; set; }
        public bool OrderedListUsesSequence { get; set; }
        public bool UseIndentsForCodeBlocks { get; set; }
        public bool PrettyPipeTables { get; set; }
        public bool IgnoreMetadata { get; set; }
        public bool UseGenericIds { get; set; }

        public override string ToString()
        {
            return $"IsArticle:{IsArticle}, Document:{Document}";
        }
    }
}
