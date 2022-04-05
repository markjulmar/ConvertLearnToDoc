using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using ConvertLearnToDoc.AzureFunctions.Models;
using LearnDocUtils;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace ConvertLearnToDoc.AzureFunctions;

public static class PageToDoc
{
    [FunctionName("PageToDoc")]
    public static async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
    {
        log.LogInformation("PageToDoc invoked.");

        string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
        var model = JsonConvert.DeserializeObject<PageToDocModel>(requestBody);
        if (model == null || !model.IsValid())
        {
            log.LogError("Bad model received.");
            return new BadRequestErrorMessageResult("Invalid request.");
        }

        model.Folder = model.Folder.Replace('\\', '/');
        if (!model.Folder.StartsWith('/'))
            model.Folder = '/' + model.Folder;

        bool isLocal = Environment.GetEnvironmentVariable("AZURE_FUNCTIONS_ENVIRONMENT") == "Development";
        var gitHubToken = Environment.GetEnvironmentVariable("GitHubToken");
        if (string.IsNullOrEmpty(gitHubToken) && !isLocal)
        {
            log.LogError("Missing GitHubToken in Function environment.");
            return new BadRequestResult();
        }

        var outputFile = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(
            model.Folder.Split('/').Last(s => !string.IsNullOrWhiteSpace(s)), "docx"));

        try
        {
            log.LogDebug($"PageToDoc(repo:{model.Repository}, branch:{model.Branch}, folder:{model.Folder}: outputFile={outputFile})");
            await SinglePageToDocx.ConvertFromRepoAsync(MSLearnRepos.Constants.DocsOrganization, model.Repository, model.Branch, model.Folder, outputFile,
                gitHubToken, new DocumentOptions { ZonePivot = model.ZonePivot });
        }
        catch (Exception ex)
        {
            log.LogError(ex.ToString());
            return new BadRequestErrorMessageResult($"Error: {ex.Message}");
        }

        if (File.Exists(outputFile))
        {
            try
            {
                return new FileContentResult(await File.ReadAllBytesAsync(outputFile), Constants.WordMimeType)
                { FileDownloadName = Path.GetFileName(outputFile) };
            }
            finally
            {
                File.Delete(outputFile);
            }
        }

        return new BadRequestErrorMessageResult(
            $"Unable to convert {model.Repository}:{model.Branch}/{model.Folder} to a Word document.");
    }
}