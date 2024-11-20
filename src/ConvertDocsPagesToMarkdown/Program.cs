using Julmar.DocsToMarkdown;

string[] testUrls =
[
    //@"https://learn.microsoft.com/en-us/java/openjdk/overview",
    //@"https://learn.microsoft.com/en-us/azure/app-service/overview-security",
    //@"https://learn.microsoft.com/en-us/azure/app-service/monitor-app-service",
    //@"https://learn.microsoft.com/en-us/azure/app-service/configure-linux-open-ssh-session?pivots=container-linux",
    //@"https://learn.microsoft.com/en-us/azure/app-service/monitor-instances-health-check?tabs=dotnet",
    //@"https://learn.microsoft.com/en-us/training/modules/secure-ai-services/",
    //@"https://learn.microsoft.com/en-us/training/modules/secure-ai-services/5-knowledge-check",
    //@"https://learn.microsoft.com/en-us/training/modules/get-started-openai"
    @"https://learn.microsoft.com/en-us/training/modules/describe-cloud-service-types/"
];

var rootFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

foreach (var url in testUrls)
{
    var converter = new DocsConverter(rootFolder, new Uri(url, UriKind.Absolute));
    var files = await converter.ConvertAsync();
    foreach (var file in files)
    {
        Console.WriteLine(file);
    }
}
