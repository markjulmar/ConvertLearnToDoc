using System.Net;
using System.Net.Sockets;
using System.Text;
using Julmar.DocsToMarkdown;

namespace ConvertLearnToDoc.UnitTests;

public class RedirectedPageTests
{
    [Theory]
    [InlineData("/guide/0-start/license/")]
    [InlineData("/old/location")]
    [InlineData("/guide/0-start/license")]
    public async Task RelativeImagesUseTheFinalPageUrl(string requestedPath)
    {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        using var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var origin = new Uri($"http://127.0.0.1:{((IPEndPoint)listener.LocalEndpoint).Port}");
        var pageUrl = new Uri(origin, "/guide/0-start/license");
        var server = ServePageAsync(listener, timeout.Token);

        try
        {
            var assets = new List<Uri>();
            var converter = await IHtmlConverter.LoadAsync(new Uri(origin, requestedPath),
                new HtmlConverterOptions
                {
                    UseDocfxExtensions = true,
                    DownloadAsset = (_, url) => assets.Add(url)
                });
            var markdown = string.Join("\n", converter.GetMarkdown());

            Assert.Equal(pageUrl, converter.Url);
            Assert.Equal(new Uri(origin, "/guide/images/icon.png"), Assert.Single(assets));
            Assert.Contains(":::image type=\"icon\"", markdown);
        }
        finally
        {
            timeout.Cancel();
            listener.Stop();
            try
            {
                await server;
            }
            catch (OperationCanceledException) { }
        }
    }

    private static async Task ServePageAsync(TcpListener listener, CancellationToken cancellationToken)
    {
        const string finalPath = "/guide/0-start/license";
        const string html = """
            <html><head><meta name="page_type" content="conceptual"></head>
            <body><main><h1>License</h1><p><img src="../images/icon.png" role="presentation"></p></main></body></html>
            """;

        while (true)
        {
            using var connection = await listener.AcceptTcpClientAsync(cancellationToken);
            await using var stream = connection.GetStream();
            using var reader = new StreamReader(stream, Encoding.ASCII, leaveOpen: true);
            var request = await reader.ReadLineAsync(cancellationToken);
            while (!string.IsNullOrEmpty(await reader.ReadLineAsync(cancellationToken))) { }

            bool isFinalPage = request?.Split(' ')[1] == finalPath;
            var body = isFinalPage ? Encoding.UTF8.GetBytes(html) : [];
            var headers = isFinalPage
                ? "HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\n"
                : $"HTTP/1.1 302 Found\r\nLocation: {finalPath}\r\n";
            headers += $"Content-Length: {body.Length}\r\nConnection: close\r\n\r\n";
            await stream.WriteAsync(Encoding.ASCII.GetBytes(headers), cancellationToken);
            await stream.WriteAsync(body, cancellationToken);

            if (isFinalPage)
                return;
        }
    }
}
