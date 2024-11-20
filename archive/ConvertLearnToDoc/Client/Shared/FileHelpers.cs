using Microsoft.JSInterop;

namespace ConvertLearnToDoc.Client.Shared
{
    public static class FileHelpers
    {
        public static string GetFilenameFromResponse(HttpResponseMessage response, string defaultFileName)
        {
            if (!response.Content.Headers.TryGetValues("Content-Disposition", out var values)) 
                return defaultFileName;
            
            const string filenameTag = "filename=";
            foreach (var cd in values)
            {
                int pos = cd.IndexOf(filenameTag, StringComparison.InvariantCultureIgnoreCase);
                if (pos > 0)
                {
                    pos += filenameTag.Length;
                    defaultFileName = cd[pos..];
                }
            }

            return defaultFileName;
        }

        public static async Task DownloadFileFromResponseAsync(HttpResponseMessage response, IJSRuntime jsRuntime, string fileName)
        {
            using var streamRef = new DotNetStreamReference(await response.Content.ReadAsStreamAsync(), leaveOpen: false);
            await jsRuntime.InvokeVoidAsync("saveFile", fileName, streamRef);
        }
    }
}
