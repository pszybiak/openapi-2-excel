namespace openapi2excel;

public class OpenApiDocumentationOptions
{
    private OpenApiDocumentationLanguage? _language = null;

    public OpenApiDocumentationLanguage Language
    {
        get => _language ?? OpenApiDocumentationLanguage.Default;
        set => _language = value;
    }
}

public class OpenApiDocumentationLanguage : Dictionary<string, string>
{
    public string Get(string key)
    {
        return TryGetValue(key, out var result) ? result : "<LANG_ERROR>";
    }

    internal static OpenApiDocumentationLanguage Default
        => new()
        {
            { OpenApiDocumentationLanguageConst.Info, "Info" },
            { OpenApiDocumentationLanguageConst.Title, "Title" },
            { OpenApiDocumentationLanguageConst.Version, "Version" },
            { OpenApiDocumentationLanguageConst.Description, "Description" },
        };
}

public static class OpenApiDocumentationLanguageConst
{
    public const string Info = "Info";
    public const string Title = "Title";
    public const string Version = "Version";
    public const string Description = "Description";
}