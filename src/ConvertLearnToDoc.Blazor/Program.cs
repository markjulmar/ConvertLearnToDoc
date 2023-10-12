using ConvertLearnToDoc.Utility;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication.MicrosoftAccount;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Server;
using Microsoft.Extensions.Options;
using System.Net;
using System.Security.Claims;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddScoped<AuthenticationStateProvider, ServerAuthenticationStateProvider>();
builder.Services.AddApplicationInsightsTelemetry();
builder.Services.AddHttpClient();

builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie()
    .AddMicrosoftAccount(ms => {
        ms.ClientId = builder.Configuration["AzureApp:ClientId"]!;
        ms.ClientSecret = builder.Configuration["AzureApp:SecretKey"]!;
        ms.SaveTokens = true;
    });

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();

/*
app.UseCookiePolicy(new CookiePolicyOptions
{
    MinimumSameSitePolicy = SameSiteMode.None
});
*/

app.UseAuthentication();
app.UseAuthorization();

app.MapPost("/signout", async ctx =>
{
    await ctx.SignOutAsync(new AuthenticationProperties { RedirectUri = "/" });
});

app.MapGet("/signin", async ctx =>
{
    await ctx.ChallengeAsync(MicrosoftAccountDefaults.AuthenticationScheme, 
        new AuthenticationProperties { RedirectUri = "/signin-callback" });
});

app.MapGet("/signin-callback", async ctx =>
{
    var result = await ctx.AuthenticateAsync();
    if (result.Succeeded)
    {
        var user = result.Principal;
        var email = user?.Claims.FirstOrDefault(c => c.Type == ClaimTypes.Email)?.Value;
        bool isValid = ControllerExtensions.ValidateUser(email);

        if (isValid && user?.Identity?.IsAuthenticated == true && !string.IsNullOrEmpty(email))
        {
            //var claimsIdentity = new ClaimsIdentity(new[] { 
            //    new Claim(ClaimTypes.Name, user.Identity.Name ?? "User"),
            //    new Claim(ClaimTypes.Email, email)
            //}, CookieAuthenticationDefaults.AuthenticationScheme);

            //var claimsPrincipal = new ClaimsPrincipal(claimsIdentity);
            await ctx.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, user);

            ctx.Response.Redirect("/");
        }
        else
        {
            await ctx.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme,
                    new AuthenticationProperties
                    {
                        RedirectUri = "/"
                    });
        }
    }
    else
    {
        await ctx.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme,
                new AuthenticationProperties
                {
                    RedirectUri = "/"
                });
    }
});

app.MapControllers();
app.MapBlazorHub();
app.MapFallbackToPage("/_Host");

app.Run();
