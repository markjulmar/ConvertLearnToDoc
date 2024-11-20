using System;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using LearnDocUtils;

namespace ConvertLearnToDoc.Tests.Int
{
    internal static class TestUtilities
    {
        public static string CreateTestModuleFolder()
        {
            string folder = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(folder);
            ZipFile.ExtractToDirectory("sample-module.zip", folder, overwriteFiles: true);
            return folder;
        }

        public static string CreateTestWordDoc()
        {
            string wordDoc = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetRandomFileName(), "docx"));
            File.Copy("sample.docx", wordDoc);
            return wordDoc;
        }

        public static Task ConvertLearnModuleToWordDocumentAsync(string inputFolder, string outputDoc)
        {
            if (inputFolder == null) throw new ArgumentNullException(nameof(inputFolder));
            if (outputDoc == null) throw new ArgumentNullException(nameof(outputDoc));
            if (!Directory.Exists(inputFolder)) throw new ArgumentException($"{inputFolder} does not exist.");
            
            if (File.Exists(outputDoc)) 
                File.Delete(outputDoc);

            return LearnToDocx.ConvertFromFolderAsync("", inputFolder, outputDoc);

        }

        public static Task ConvertWordDocumentToLearnFolderAsync(string inputDoc, string outputFolder)
        {
            if (!File.Exists(inputDoc)) throw new ArgumentException($"{inputDoc} does not exist.");
            if (outputFolder == null) throw new ArgumentNullException(nameof(outputFolder));

            if (Directory.Exists(outputFolder))
                Directory.Delete(outputFolder, true);

            return DocxToLearn.ConvertAsync(inputDoc, outputFolder, new MarkdownOptions());
        }
    }
}
