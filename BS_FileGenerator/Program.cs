using BS_FileGenerator;
using BS_FileGenerator.Middleware;

var builder = WebApplication.CreateBuilder(args);

var startup = new Startup(builder);
startup.ConfiguraServices();

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
