namespace LearnDocUtils;

static class ConvertUrls
{
    public static string FromRelative(string url, string baseUrl)
    {
        if (string.IsNullOrWhiteSpace(url))
            return string.Empty;

        // Skip if not a full URL, or already an absolute URL.
        if (!Uri.TryCreate(url, UriKind.RelativeOrAbsolute, out var uri) || uri.IsAbsoluteUri)
            return url;
        
        if (!string.IsNullOrWhiteSpace(baseUrl) 
            && Uri.TryCreate(baseUrl, UriKind.Absolute, out var baseUri) 
            && Uri.TryCreate(baseUri, uri, out var combinedUri)) 
            return combinedUri.ToString();

        return url;
    }

    public static string FromAbsolute(string url, string baseUrl)
    {
        if (string.IsNullOrWhiteSpace(url))
            return string.Empty;

        // Skip if not a full URL.
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || !uri.IsAbsoluteUri)
            return url;

        // Step 1: check to see if this document is somewhere along the path of the original URL we
        // created this from.
        if (!string.IsNullOrWhiteSpace(baseUrl) && Uri.TryCreate(baseUrl, UriKind.Absolute, out var baseUri))
        {
            if (baseUri.Host.Equals(uri.Host, StringComparison.CurrentCultureIgnoreCase))
            {
                var baseSegments = baseUri.Segments;
                var uriSegments = uri.Segments;
                if (baseSegments.Length > 1 && uriSegments.Length > 1)
                {
                    var commonSegments = baseSegments.TakeWhile((t, i) 
                        => i < uriSegments.Length 
                           && t.Equals(uriSegments[i], StringComparison.OrdinalIgnoreCase)).Count();
                    
                    if (commonSegments > 1)
                    {
                        var backup = baseSegments.Length - commonSegments;
                        var relativeUrl = backup > 1 ? string.Join("/", Enumerable.Range(0, baseSegments.Length-commonSegments).Select(_ => "..")) : "";
                        if (!string.IsNullOrEmpty(relativeUrl)) relativeUrl += '/';
                        relativeUrl += string.Join("", uriSegments[commonSegments..]) + ".md";
                        return relativeUrl;
                    }
                }
            }
        }

        // Step 2: see if this is a URL on the Learn site (or review) and if so, convert it to a relative URL.
        if (uri.Host.Contains(Constants.LearnSiteHostName, StringComparison.CurrentCultureIgnoreCase))
        {
            var segments = uri.Segments;
            if (segments.Length <= 2 || segments[1].Length != 6 || !segments[1].Contains('-'))
                return uri.AbsolutePath;
                
            // Remove any locale
            var relativeUrl = string.Join("", segments[2..]);
            return !relativeUrl.StartsWith('/') ? "/" + relativeUrl : relativeUrl;
        }

        return url;
    }
}