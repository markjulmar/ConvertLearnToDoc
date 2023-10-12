using Microsoft.AspNetCore.Mvc;
using System.Security.Claims;

namespace ConvertLearnToDoc.Utility;

public static class ControllerExtensions
{
    public static string AnonymousIdentity = "Anonymous";

    public static string GetUsername(HttpContext context)
    {
        var user = context.User;
        return user?.Identity?.IsAuthenticated != true
            ? AnonymousIdentity
            : user.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Email)?.Value ?? AnonymousIdentity;
    }

    public static bool ValidateUser(string? email)
    {
        if (string.IsNullOrWhiteSpace(email))
            return false;

        var trimmedEmail = email?.Trim() ?? "";
        if (trimmedEmail.EndsWith('.'))
            return false;

        try
        {
            var addr = new System.Net.Mail.MailAddress(email ?? "");
            return addr.Host is "microsoft.com" or "julmar.com";
        }
        catch
        {
            // Ignore
        }

        return false;
    }

    public static async Task<IActionResult> FileAttachment(this ControllerBase controller, string filename, string contentType, bool deleteFile = true)
    {
        try
        {
            var content = await File.ReadAllBytesAsync(filename);
            controller.Response.Headers
                .Add("Content-Disposition",
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