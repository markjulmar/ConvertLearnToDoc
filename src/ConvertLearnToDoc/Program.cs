using ConvertLearnToDoc.Utility;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
#if USE_MS_AUTH
using Microsoft.AspNetCore.Authentication.MicrosoftAccount;
#else
using Microsoft.AspNetCore.Authentication.OAuth;
#endif
using Microsoft.AspNetCore.Components.Authorization;
using Microsoft.AspNetCore.Components.Server;
using System.Security.Claims;
using Microsoft.Extensions.Logging.ApplicationInsights;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddScoped<AuthenticationStateProvider, ServerAuthenticationStateProvider>();
#if !DEBUG
builder.Services.AddApplicationInsightsTelemetry();
#endif

builder.Services.AddHttpClient();
builder.Services.AddHttpContextAccessor();
#if !USE_MS_AUTH
builder.Services.AddScoped<TokenService>();
#endif

builder.Logging.SetMinimumLevel(LogLevel.Debug);
builder.Logging.AddConsole();
#if !DEBUG
builder.Logging.AddApplicationInsights();
builder.Logging.AddFilter<ApplicationInsightsLoggerProvider>("", LogLevel.Warning);
#endif

#if USE_MS_AUTH
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie()
    .AddMicrosoftAccount(ms => {
        ms.ClientId = builder.Configuration["AzureApp:ClientId"]!;
        ms.ClientSecret = builder.Configuration["AzureApp:SecretKey"]!;
        ms.SaveTokens = true;
    });
#else
var gitHubAppInfo = builder.Configuration.GetSection("GitHub") 
    ?? throw new Exception("GitHub section is missing in the configuration file.");

builder.Services.AddAuthentication(options =>
{
    options.DefaultAuthenticateScheme = CookieAuthenticationDefaults.AuthenticationScheme;
    options.DefaultSignInScheme = CookieAuthenticationDefaults.AuthenticationScheme;
    options.DefaultChallengeScheme = "GitHub";
})
.AddCookie(options =>
{
    options.ExpireTimeSpan = TimeSpan.FromMinutes(60); // Default cookie expiration time
    options.Events = new CookieAuthenticationEvents
    {
        OnSigningIn = context =>
        {
            var expiresAt = context.Properties.GetTokenValue("expires_at");
            if (DateTime.TryParse(expiresAt, out var expiry))
            {
                context.Properties.ExpiresUtc = expiry;
                context.CookieOptions.Expires = expiry;
            }
            return Task.CompletedTask;
        }
    };
})
.AddOAuth("GitHub", options =>
{
    options.ClientId = gitHubAppInfo["ClientId"] ?? throw new Exception("GitHub section missing ClientId.");;
    options.ClientSecret = gitHubAppInfo["ClientSecret"] ?? throw new Exception("GitHub section missing ClientSecret.");
    options.CallbackPath = PathString.FromUriComponent(new Uri(gitHubAppInfo["CallbackUrl"] ?? throw new Exception("GitHub section missing CallbackUrl.")));

    options.AuthorizationEndpoint = "https://github.com/login/oauth/authorize";
    options.TokenEndpoint = "https://github.com/login/oauth/access_token";
    options.UserInformationEndpoint = "https://api.github.com/user";

    options.Scope.Add("read:user");
#if USE_GITHUB_PAT
    options.Scope.Add("repo");
    options.Scope.Add("read:org");
#endif

    options.ClaimActions.MapJsonKey(ClaimTypes.NameIdentifier, "id");
    options.ClaimActions.MapJsonKey(ClaimTypes.Name, "name");
    options.ClaimActions.MapJsonKey(ClaimTypes.Email, "email");

    options.SaveTokens = true;

    options.Events = new OAuthEvents
    {
        OnCreatingTicket = async context =>
        {
            var request = new HttpRequestMessage(HttpMethod.Get, context.Options.UserInformationEndpoint);
            request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", context.AccessToken);

            var response = await context.Backchannel.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, context.HttpContext.RequestAborted);
            response.EnsureSuccessStatusCode();

            var user = System.Text.Json.JsonDocument.Parse(await response.Content.ReadAsStringAsync());

            context.RunClaimActions(user.RootElement);

            // Save the access token
            if (context.AccessToken != null)
            {
                context.Identity!.AddClaim(new Claim("access_token", context.AccessToken));
                var expiresAt = DateTime.UtcNow.AddSeconds(int.Parse(context.TokenResponse.ExpiresIn ?? "3600"));
                context.Properties.StoreTokens(
                [
                    new AuthenticationToken { Name = "access_token", Value = context.AccessToken },
                    new AuthenticationToken { Name = "expires_at", Value = expiresAt.ToString("o") }
                ]);
                context.Properties.ExpiresUtc = expiresAt;
            }
        }
    };
});
#endif

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
#if USE_MS_AUTH
    await ctx.ChallengeAsync(MicrosoftAccountDefaults.AuthenticationScheme, 
        new AuthenticationProperties { RedirectUri = "/signin-callback" });
#else
    await ctx.ChallengeAsync("GitHub", new AuthenticationProperties { RedirectUri = "/signin-callback" });
#endif
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
