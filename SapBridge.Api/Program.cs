using Serilog;
using SapBridge.Core;

var builder = WebApplication.CreateBuilder(args);

// Configure Serilog
Log.Logger = new LoggerConfiguration()
    .WriteTo.Console()
    .WriteTo.File("logs/sap-bridge.log", rollingInterval: RollingInterval.Day)
    .CreateLogger();

builder.Host.UseSerilog();

// Add services
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Add CORS
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

// Register application services
builder.Services.AddSingleton(Log.Logger);
builder.Services.AddSingleton<SapGuiConnector>();
builder.Services.AddSingleton<ComIntrospector>(); 
builder.Services.AddSingleton<ActionExecutor>();
builder.Services.AddSingleton<ScreenService>();

var app = builder.Build();

// Configure middleware
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseCors();
app.UseAuthorization();
app.MapControllers();

Log.Information("SAP Bridge starting on http://0.0.0.0:5000");
app.Run("http://0.0.0.0:5000");

