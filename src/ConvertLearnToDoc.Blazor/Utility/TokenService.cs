using Microsoft.AspNetCore.Authentication;

public class TokenService(IHttpContextAccessor httpContextAccessor)
{
    public async Task<string?> GetAccessTokenAsync()
    {
        if (httpContextAccessor.HttpContext != null)
        {
            var authenticateResult = await httpContextAccessor.HttpContext.AuthenticateAsync();
            if (authenticateResult.Succeeded)
            {
                var token = authenticateResult.Principal?.FindFirst("access_token")?.Value;
                return token;
            }
        }
        return null;
    }
}
