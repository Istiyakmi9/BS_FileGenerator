
using Bot.CoreBottomHalf.CommonModal;
using Bot.CoreBottomHalf.CommonModal.Enums;
using BS_FileGenerator.IService;
using BS_FileGenerator.Middleware;
using BS_FileGenerator.Models;
using BS_FileGenerator.Service;
using Confluent.Kafka;
using DinkToPdf;
using DinkToPdf.Contracts;
using DocMaker.HtmlToDocx;
using Microsoft.Extensions.Options;
using ModalLayer;
using Newtonsoft.Json.Serialization;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Configuration.SetBasePath(builder.Environment.ContentRootPath)
    .AddJsonFile($"appsettings.{builder.Environment.EnvironmentName}.json", optional: false, reloadOnChange: true)
    .AddEnvironmentVariables();
builder.Services.AddSingleton<ApplicationConfiguration>();

var kafkaServerDetail = new ProducerConfig();
builder.Configuration.Bind("KafkaServerDetail", kafkaServerDetail);
builder.Services.Configure<KafkaServiceConfig>(x => builder.Configuration.GetSection(nameof(KafkaServiceConfig)).Bind(x));
builder.Services.Configure<JwtSetting>(o => builder.Configuration.GetSection(nameof(JwtSetting)).Bind(o));

builder.Services.AddSingleton<ProducerConfig>(kafkaServerDetail);
builder.Services.AddSingleton<KafkaNotificationService>(x =>
{
    return new KafkaNotificationService(
        x.GetRequiredService<IOptions<KafkaServiceConfig>>(),
        x.GetRequiredService<ProducerConfig>(),
        x.GetRequiredService<ILogger<KafkaNotificationService>>(),
        builder.Environment.EnvironmentName == nameof(DefinedEnvironments.Development) ?
                        DefinedEnvironments.Development :
                        DefinedEnvironments.Production
    );
});

builder.Services.AddScoped<CurrentSession>(x =>
{
    return new CurrentSession
    {
        Environment = builder.Environment.EnvironmentName == nameof(DefinedEnvironments.Development) ?
                        DefinedEnvironments.Development :
                        DefinedEnvironments.Production
    };
});

builder.Services.AddControllers().AddNewtonsoftJson(options =>
{
    options.SerializerSettings.ReferenceLoopHandling = Newtonsoft.Json.ReferenceLoopHandling.Ignore;
    options.SerializerSettings.ContractResolver = new DefaultContractResolver();
});
builder.Services.AddScoped<IHtmlToPdfConverter, HtmlToPdfConverter>();
builder.Services.AddScoped<IHtmlToDocxConverter, HtmlToDocxConverter>();
builder.Services.AddScoped<IDataTableToExcel, DataTableToExcel>();
builder.Services.AddScoped<IDataTableToExcel, DataTableToExcel>();
builder.Services.AddSingleton(typeof(IConverter), new SynchronizedConverter(new PdfTools()));
builder.Services.AddScoped<IHTMLConverter, HTMLConverter>();

builder.Services.AddCors(options =>
{
    options.AddPolicy("bottomhalf-cors", policy =>
    {
        policy.AllowAnyOrigin()
        .AllowAnyHeader()
        .AllowAnyMethod()
        .WithExposedHeaders("Authorization");
    });
});

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}
app.UseMiddleware<ExceptionHandlerMiddleware>();
app.UseCors("bottomhalf-cors");
app.UseAuthorization();

app.MapControllers();

app.Run();
