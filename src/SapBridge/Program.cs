using Serilog;
using SapBridge.Repositories;
using SapBridge.Services.Grid;
using SapBridge.Services.Table;
using SapBridge.Services.Tree;
using SapBridge.Services.Vision;
using SapBridge.Services.Query;
using SapBridge.Services.Session;
using SapBridge.Services.Screen;

var builder = WebApplication.CreateBuilder(args);

// Configure Serilog
Log.Logger = new LoggerConfiguration()
    .ReadFrom.Configuration(builder.Configuration)
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .WriteTo.File("logs/sap-bridge-.log", rollingInterval: RollingInterval.Day)
    .CreateLogger();

builder.Host.UseSerilog();

// Add services to the container
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// Register Serilog
builder.Services.AddSingleton(Log.Logger);

// Register repositories
builder.Services.AddSingleton<ISapGuiRepository, SapGuiRepository>();

// Register core services
builder.Services.AddScoped<ISessionService, SessionService>();
builder.Services.AddScoped<IScreenService, ScreenService>();

// Register query services
builder.Services.AddScoped<QueryValidator>();
builder.Services.AddScoped<ConditionEvaluator>();
builder.Services.AddScoped<IQueryEngine, QueryEngine>();

// Register data services
builder.Services.AddScoped<IGridService, GridService>();
builder.Services.AddScoped<ITableService, TableService>();
builder.Services.AddScoped<ITreeService, TreeService>();

// Register vision service
builder.Services.AddScoped<IVisionService, VisionService>();

var app = builder.Build();

// Configure the HTTP request pipeline
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseSerilogRequestLogging();

app.UseAuthorization();

app.MapControllers();

// Log startup
Log.Information("SAP Bridge starting up...");
Log.Information("Environment: {Environment}", app.Environment.EnvironmentName);

try
{
    app.Run();
}
catch (Exception ex)
{
    Log.Fatal(ex, "Application terminated unexpectedly");
}
finally
{
    Log.CloseAndFlush();
}

