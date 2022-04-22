using AE.O365.GetFiles.CSApp.Common.Interfaces;
using AE.O365.GetFiles.CSApp.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace AE.O365.GetFiles.CSApp;

public class Program
{
    static void Main(string[] args)
    {
        var serviceCollection = new ServiceCollection();
        ConfigureServices(serviceCollection);

        var serviceProvider = serviceCollection.BuildServiceProvider();

        var logger = serviceProvider
            .GetRequiredService<ILoggerFactory>()
            .CreateLogger<Program>();

        try
        {

            var getfiles = serviceProvider.GetService<IService>();
            getfiles.Run();

            logger.LogDebug("All done!");
        }
        catch (Exception ex)
        {
            logger.LogError(ex, ex.Message);
        }
    }

    static void ConfigureServices(IServiceCollection serviceCollection)
    {
        serviceCollection.AddSingleton<ILoggerFactory, LoggerFactory>();
        serviceCollection.AddSingleton(typeof(ILogger<>), typeof(Logger<>));
        serviceCollection.AddLogging(loggingBuilder => loggingBuilder
            .AddConsole()
            .AddDebug()
            .SetMinimumLevel(LogLevel.Debug));

        var configuration = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", false)
            .Build();

        serviceCollection.AddSingleton(configuration);

        serviceCollection.AddTransient<IService, GetFileService>();
    }
}