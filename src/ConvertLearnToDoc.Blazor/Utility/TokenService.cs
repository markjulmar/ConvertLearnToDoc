using Microsoft.AspNetCore.Authentication;

public class TokenService
{
    private readonly IHttpContextAccessor _httpContextAccessor;

    public TokenService(IHttpContextAccessor httpContextAccessor)
    {
        _httpContextAccessor = httpContextAccessor;
    }

    public async Task<string?> GetAccessTokenAsync()
    {
        var authenticateResult = await _httpContextAccessor.HttpContext.AuthenticateAsync();
        if (authenticateResult.Succeeded)
        {
            var token = authenticateResult.Principal?.FindFirst("access_token")?.Value;
            return token;
        }
        return null;
    }
}
