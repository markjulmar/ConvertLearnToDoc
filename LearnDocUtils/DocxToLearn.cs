using System;
using System.IO;
using System.Threading.Tasks;

namespace LearnDocUtils
{
    public static class DocxToLearn
    {
        public static Task ConvertAsync(string docxFile, string outputFolder,
            Action<string> logger, bool debug, bool usePandoc)
        {
            if (string.IsNullOrWhiteSpace(docxFile))
                throw new ArgumentException($"'{nameof(docxFile)}' cannot be null or whitespace.", nameof(docxFile));
            if (!File.Exists(docxFile))
                throw new ArgumentException($"Error: {docxFile} does not exist.", nameof(docxFile));
            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new ArgumentException($"'{nameof(outputFolder)}' cannot be null or whitespace.", nameof(outputFolder));

            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            if (usePandoc)
            {
                return new DocxToLearnPandoc().ConvertAsync(docxFile, outputFolder, logger, debug);
            }

            return new DocxToLearnDXPlus().ConvertAsync(docxFile, outputFolder, logger, debug);
        }
    }
}