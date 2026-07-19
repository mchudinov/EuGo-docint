using DocInt.Api.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;

namespace DocInt.Tests;

public class OptionsTests
{
    [Fact]
    public void Defaults_match_the_spec()
    {
        var o = new DocIntOptions();
        Assert.Equal(52_428_800, o.MaxFileBytes);
        Assert.Equal(32, o.MaxFilesPerRequest);
        Assert.Equal(100, o.PerFileTimeoutSeconds);
        Assert.Equal(4, o.MaxParallelism);
        Assert.Equal(52_428_800L * 32 + 1_048_576, o.MaxRequestBytes);
        Assert.Equal("gpt-4.1-mini", new AzureOpenAIOptions().DeploymentNameVision);
    }

    [Fact]
    public void Configuration_binds_into_options()
    {
        using var factory = new BindingFactory();
        var docint = factory.Services.GetRequiredService<IOptions<DocIntOptions>>().Value;
        var di = factory.Services.GetRequiredService<IOptions<DocumentIntelligenceOptions>>().Value;
        Assert.Equal(3, docint.MaxFilesPerRequest);
        Assert.Equal("https://di.example", di.Endpoint);
    }

    private sealed class BindingFactory : DocIntAppFactory
    {
        protected override void ConfigureWebHost(Microsoft.AspNetCore.Hosting.IWebHostBuilder builder)
        {
            builder.UseSetting("DocInt:MaxFilesPerRequest", "3");
            builder.UseSetting("DocumentIntelligence:Endpoint", "https://di.example");
            base.ConfigureWebHost(builder);
        }
    }
}
