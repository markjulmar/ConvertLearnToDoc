using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using ConvertLearnToDoc.AzureFunctions;
using ConvertLearnToDoc.AzureFunctions.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System.Text.Json;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Internal;
using Microsoft.Extensions.Primitives;

namespace ConvertLearnToDoc.Tests.Int
{
    internal static class TestUtilities
    {
        public static async Task<byte[]> CreateWordDoc(LearnToDocModel model, ILogger logger)
        {
            var request = TestFactory.CreateHttpRequest("", "");

            request.Method = "POST";
            var json = JsonSerializer.Serialize(model);
            byte[] bytes = Encoding.UTF8.GetBytes(json);
            MemoryStream stream = new(bytes);
            request.Body = stream;
            request.ContentLength = bytes.Length;
            request.ContentType = "application/x-www-form-urlencoded";

            var fileContentResult = (FileContentResult)await LearnToDoc.Run(request, logger);
            return fileContentResult.FileContents;
        }

        public static async Task<byte[]> ConvertDocxToModuleZipFile(string wordDocFilePath, string fileName, ILogger logger)
        {
            var request = TestFactory.CreateHttpRequest("", "");
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";

            var fileCollection = new FormFileCollection();
            var bytes = await File.ReadAllBytesAsync(wordDocFilePath);
            MemoryStream stream = new(bytes);

            var wordFile = new FormFile(stream, 0, bytes.Length, "wordDoc", fileName)
            {
                Headers = new HeaderDictionary
                {
                    new("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                }
            };
            fileCollection.Add(wordFile);

            const bool UseAsterisksForBullets = false;
            const bool UseAsterisksForEmphasis = true;
            const bool OrderedListUsesSequence = false;
            const bool UseAlternateHeaderSyntax = false;
            const bool UseIndentsForCodeBlocks = false;

            var formCollection = new FormCollection(new Dictionary<string, StringValues>
            {
                { nameof(UseAsterisksForBullets), UseAsterisksForBullets.ToString() },
                { nameof(UseAsterisksForEmphasis), UseAsterisksForEmphasis.ToString() },
                { nameof(OrderedListUsesSequence), OrderedListUsesSequence.ToString() },
                { nameof(UseAlternateHeaderSyntax), UseAlternateHeaderSyntax.ToString() },
                { nameof(UseIndentsForCodeBlocks), UseIndentsForCodeBlocks.ToString() },
            }, fileCollection);

            request.Form = formCollection;

            var fileContentResult = (FileContentResult)await DocToLearn.Run(request, logger);

            return fileContentResult.FileContents;
        }

    }
}
