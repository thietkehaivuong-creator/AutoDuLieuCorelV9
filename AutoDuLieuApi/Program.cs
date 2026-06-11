var builder = WebApplicationBuilder.CreateBuilder(args);

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new Microsoft.OpenApi.Models.OpenApiInfo
    {
        Title = "Auto Du Lieu Corel V9 API",
        Version = "v1",
        Description = "API cho hệ thống tự động hóa dữ liệu CorelDRAW"
    });
});

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

builder.Services.AddScoped<IDataProcessingService, DataProcessingService>();
builder.Services.AddScoped<ICredentialsService, CredentialsService>();
builder.Services.AddLogging();

builder.Configuration
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .AddEnvironmentVariables();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI(c =>
    {
        c.SwaggerEndpoint("/swagger/v1/swagger.json", "Auto Du Lieu API v1");
        c.RoutePrefix = "swagger";
    });
}

app.UseHttpsRedirection();
app.UseCors("AllowAll");
app.UseRouting();
app.MapControllers();

app.MapGet("/", () =>
{
    return Results.Ok(new
    {
        name = "Auto Du Lieu Corel V9 API",
        version = "1.0.0",
        swagger = "/swagger",
        endpoints = new
        {
            search = "GET /api/data/search/{pfi}",
            all = "GET /api/data/all",
            formatDate = "POST /api/data/format-date",
            settings = "GET /api/settings/status",
            saveCredentials = "POST /api/settings/credentials"
        }
    });
});

app.Run();