using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using WebApplication1.Command;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddTransient<PostDocumentCommand>();
// Register MVC controllers + Swagger (optional)
builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Dev‐only diagnostics + Swagger UI
if (app.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

// Map all controllers in /Controllers
app.MapControllers();

app.Run();

