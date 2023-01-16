﻿using Microsoft.AspNetCore.Mvc;

namespace ConvertLearnToDoc.Server.Controllers;

public static class ControllerExtensions
{
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