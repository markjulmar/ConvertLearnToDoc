using Microsoft.AspNetCore.Mvc;

namespace ConvertLearnToDoc.Server.Controllers;

public static class ControllerExtensions
{
    public static string? GetUsername(HttpContext context)
    {
        const string identifier = "X-MS-CLIENT-PRINCIPAL-NAME";
        return context.Request.Headers[identifier].FirstOrDefault();
    }

    public static async Task<IActionResult> FileAttachment(this ControllerBase controller, string filename, string contentType, bool deleteFile = true)
    {
        try
        {
            var content = await File.ReadAllBytesAsync(filename);
            controller.Response.Headers
                .Append("Content-Disposition", 
                    $"attachment;filename={Path.GetFileName(filename).Replace(' ', '-')}");
            return controller.File(content, contentType);
        }
        finally
        {
            if (deleteFile)
            {
                File.Delete(filename);
            }
        }
    }
}