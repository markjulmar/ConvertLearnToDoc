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

        var trimmedEmail = email.Trim().ToLower();
        if (trimmedEmail.EndsWith('.'))
            return false;

        /*
        if (ValidUsers == null)
        {
            var users = Environment.GetEnvironmentVariable("VALID_USERS");
            ValidUsers = string.IsNullOrWhiteSpace(users) ? new() : users.Split(';').Select(s => s.Trim().ToLower()).ToList();
        }

        try
        {
            if (ValidUsers.Contains(trimmedEmail))
                return true;

            var qualifiedAddress = new System.Net.Mail.MailAddress(email);
            return qualifiedAddress.Host is "microsoft.com";
        }
        catch
        {
            // Ignore
        }

        return false;
        */

        // All are allowed.
        return true;

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