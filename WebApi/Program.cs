using Aplicacion.Services;
using Microsoft.EntityFrameworkCore;
using Persistencia.Context;
using Persistencia.Repository;
using WebAPI.Midleware;
using WebApi.Background;

var myAllowSpecificOrigins = "_myAllowSpecificOrigins";
var builder = WebApplication.CreateBuilder(args);




builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();


builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options
    .UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection"))
    .EnableSensitiveDataLogging()
    .EnableDetailedErrors()
    .UseQueryTrackingBehavior(QueryTrackingBehavior.NoTracking)
);


builder.Services.AddScoped(typeof(IGenericRepository<>), typeof(GenericRepository<>));
builder.Services.AddScoped(typeof(IGenericService<>), typeof(GenericService<>));


builder.Services.AddHostedService<ProcessData>();



builder.Services.AddCors(opt => {
    opt.AddPolicy(name: myAllowSpecificOrigins,
        builder => {
            builder.SetIsOriginAllowed(origin => new Uri(origin).Host == "localhost")
            .AllowAnyMethod()
            .AllowAnyHeader();
        });
});






var app = builder.Build();


app.UseMiddleware<ExcepcionErroresMidleware>();

using (var scope = app.Services.CreateScope())
{
    var context = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
    context.Database.Migrate();
}


if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseCors(myAllowSpecificOrigins);

app.UseAuthorization();

app.MapControllers();
app.Run();


