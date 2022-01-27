using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace LearnDocUtils
{
    static class NotebookConverter
    {
        public static async Task<string> Convert(string notebookUrl)
        {
            string jsonText = await DownloadNotebook(notebookUrl);
            if (string.IsNullOrEmpty(jsonText))
                return null;

            var sb = new StringBuilder();

            try
            {
                dynamic nb = JObject.Parse(jsonText);
                dynamic cells = (JArray) nb.cells;
                string language = (string) nb.metadata.kernelspec.language;

                foreach (dynamic cell in cells)
                {
                    if (cell.cell_type == "code")
                    {
                        FormatCode(sb, language, (JArray)cell.source);
                    }
                    else if (cell.cell_type == "markdown")
                    {
                        FormatMarkdown(sb, (JArray)cell.source);
                    }
                }
            }
            catch
            {
                return null;
            }

            return sb.ToString();
        }

        private static void FormatMarkdown(StringBuilder sb, JArray source)
        {
            foreach (var line in source)
            {
                // Change any H1 to H2 so it doesn't look like a unit.
                string text = line.ToString();
                if (text.StartsWith("# "))
                    text = "#" + text;

                sb.Append(text);
            }

            sb.AppendLine()
              .AppendLine();
        }

        private static void FormatCode(StringBuilder sb, string language, JArray source)
        {
            sb.AppendLine($"```{language}");

            foreach (var line in source)
                sb.Append(line);

            sb.AppendLine().AppendLine("```").AppendLine();
        }

        private static async Task<string> DownloadNotebook(string notebookUrl)
        {
            if (notebookUrl.ToLower().StartsWith("http"))
            {
                using var client = new HttpClient();
                return await client.GetStringAsync(notebookUrl);
            }

            return File.Exists(notebookUrl) ? await File.ReadAllTextAsync(notebookUrl) : null;
        }
    }
}
