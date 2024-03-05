namespace openapi2excel;

public class OpenApiDocumentationOptions
{
    public OpenApiDocumentationLanguage Language { get; set; } = [];
}

public class OpenApiDocumentationLanguage : Dictionary<string, string>
{
    public string Get(string key)
    {
        return TryGetValue(key, out var result) ? result : string.Empty;
    }
}

public static class OpenApiDocumentationLanguageConst
{
    public const string Info = "Info";
    public const string Title = "Title";
    public const string Version = "Version";
    public const string Description = "Description";
}