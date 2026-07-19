using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.AspNetCore.TestHost;
using Microsoft.Extensions.DependencyInjection;

namespace DocInt.Tests;

/// <summary>
/// Hosts the real Program in-memory. Subclasses override ConfigureFakes to replace
/// Azure adapters / engines; the base factory runs the app exactly as configured,
/// which means no Azure endpoints are set (the engine_unconfigured path).
/// </summary>
public class DocIntAppFactory : WebApplicationFactory<Program>
{
    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        builder.ConfigureTestServices(ConfigureFakes);
    }

    protected virtual void ConfigureFakes(IServiceCollection services)
    {
    }
}
