namespace DocInt.Api.Configuration;

public sealed class DocIntOptions
{
    public const string SectionName = "DocInt";

    public long MaxFileBytes { get; set; } = 52_428_800;
    public int MaxFilesPerRequest { get; set; } = 32;
    public int PerFileTimeoutSeconds { get; set; } = 100;
    public int MaxParallelism { get; set; } = 4;

    /// <summary>Request-level cap: worst-case accepted payload plus multipart overhead.</summary>
    public long MaxRequestBytes => MaxFileBytes * MaxFilesPerRequest + 1_048_576;
}

public sealed class DocumentIntelligenceOptions
{
    public const string SectionName = "DocumentIntelligence";

    public string? Endpoint { get; set; }
    public string? ApiKey { get; set; }
}

public sealed class AzureOpenAIOptions
{
    public const string SectionName = "AzureOpenAI";

    public string? Endpoint { get; set; }
    public string? ApiKey { get; set; }
    public string DeploymentNameVision { get; set; } = "gpt-4.1-mini";
}

public static class OptionsExtensions
{
    public static WebApplicationBuilder AddDocIntOptions(this WebApplicationBuilder builder)
    {
        builder.Services.AddOptions<DocIntOptions>()
            .Bind(builder.Configuration.GetSection(DocIntOptions.SectionName))
            .Validate(o => o.MaxFileBytes > 0 && o.MaxFilesPerRequest > 0
                        && o.PerFileTimeoutSeconds > 0 && o.MaxParallelism > 0,
                "DocInt options must all be positive")
            .ValidateOnStart();
        builder.Services.AddOptions<DocumentIntelligenceOptions>()
            .Bind(builder.Configuration.GetSection(DocumentIntelligenceOptions.SectionName))
            .Validate(o => string.IsNullOrWhiteSpace(o.Endpoint) || Uri.TryCreate(o.Endpoint, UriKind.Absolute, out _),
                $"{DocumentIntelligenceOptions.SectionName}:Endpoint must be an absolute URI")
            .ValidateOnStart();
        builder.Services.AddOptions<AzureOpenAIOptions>()
            .Bind(builder.Configuration.GetSection(AzureOpenAIOptions.SectionName))
            .Validate(o => string.IsNullOrWhiteSpace(o.Endpoint) || Uri.TryCreate(o.Endpoint, UriKind.Absolute, out _),
                $"{AzureOpenAIOptions.SectionName}:Endpoint must be an absolute URI")
            .ValidateOnStart();
        return builder;
    }
}
