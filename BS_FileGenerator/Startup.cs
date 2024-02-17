using Bot.CoreBottomHalf.CommonModal;
using Bot.CoreBottomHalf.CommonModal.Enums;
using BS_FileGenerator.IService;
using BS_FileGenerator.Models;
using BS_FileGenerator.Service;
using Confluent.Kafka;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocMaker.HtmlToDocx;
using Microsoft.Extensions.Options;
using ModalLayer;
using Newtonsoft.Json.Serialization;
using System.Reflection.Metadata.Ecma335;
using System.Runtime.Loader;

namespace BS_FileGenerator
{
    public class Startup
    {
        private readonly IConfigurationBuilder _configurationBuilder;
        private readonly IConfiguration _configuration;
        private readonly IServiceCollection _services;
        private readonly IWebHostEnvironment _environment;

        public Startup(WebApplicationBuilder builder)
        {
            _services = builder.Services;
            _configuration = builder.Configuration;
            _environment = builder.Environment;
            _configurationBuilder = builder.Configuration;
        }

        public void ConfiguraServices()
        {
            try
            {
                var context = new CustomAssemblyLoadContext();
                context.LoadUnmanagedLibrary("libwkhtmltox");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            _services.AddControllers();
            // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
            _services.AddEndpointsApiExplorer();
            _services.AddSwaggerGen();
            _configurationBuilder.SetBasePath(_environment.ContentRootPath)
                .AddJsonFile($"appsettings.{_environment.EnvironmentName}.json", optional: false, reloadOnChange: true)
                .AddEnvironmentVariables();
            _services.AddSingleton<ApplicationConfiguration>();

            var kafkaServerDetail = new ProducerConfig();
            _configuration.Bind("KafkaServerDetail", kafkaServerDetail);
            _services.Configure<KafkaServiceConfig>(x => _configuration.GetSection(nameof(KafkaServiceConfig)).Bind(x));
            _services.Configure<JwtSetting>(o => _configuration.GetSection(nameof(JwtSetting)).Bind(o));

            _services.AddSingleton<ProducerConfig>(kafkaServerDetail);
            _services.AddSingleton<KafkaNotificationService>(x =>
            {
                return new KafkaNotificationService(
                    x.GetRequiredService<IOptions<KafkaServiceConfig>>(),
                    x.GetRequiredService<ProducerConfig>(),
                    x.GetRequiredService<ILogger<KafkaNotificationService>>(),
                    _environment.EnvironmentName == nameof(DefinedEnvironments.Development) ?
                                    DefinedEnvironments.Development :
                                    DefinedEnvironments.Production
                );
            });

            _services.AddScoped<CurrentSession>(x =>
            {
                return new CurrentSession
                {
                    Environment = _environment.EnvironmentName == nameof(DefinedEnvironments.Development) ?
                                    DefinedEnvironments.Development :
                                    DefinedEnvironments.Production
                };
            });

            _services.AddControllers().AddNewtonsoftJson(options =>
            {
                options.SerializerSettings.ReferenceLoopHandling = Newtonsoft.Json.ReferenceLoopHandling.Ignore;
                options.SerializerSettings.ContractResolver = new DefaultContractResolver();
            });
            _services.AddScoped<IHtmlToPdfConverter, HtmlToPdfConverter>();
            _services.AddScoped<IHtmlToDocxConverter, HtmlToDocxConverter>();
            _services.AddScoped<IDataTableToExcel, DataTableToExcel>();
            _services.AddScoped<IDataTableToExcel, DataTableToExcel>();
            _services.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));
            _services.AddScoped<IHTMLConverter, HTMLConverter>();

            _services.AddCors(options =>
            {
                options.AddPolicy("bottomhalf-cors", policy =>
                {
                    policy.AllowAnyOrigin()
                    .AllowAnyHeader()
                    .AllowAnyMethod()
                    .WithExposedHeaders("Authorization");
                });
            });
        }
    }

    public class CustomAssemblyLoadContext : AssemblyLoadContext
    {
        public IntPtr LoadUnmanagedLibrary(string unmanagedLibraryName)
        {
            return LoadUnmanagedDll(unmanagedLibraryName);
        }

        protected override IntPtr LoadUnmanagedDll(string unmanagedDllName)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "lib", $"{unmanagedDllName}.dll");
            return LoadUnmanagedDllFromPath(path);
        }
    }
}
