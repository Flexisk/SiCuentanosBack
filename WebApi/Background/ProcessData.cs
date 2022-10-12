using Aplicacion.AgregarExcel;

namespace WebApi.Background
{
    public class ProcessData : IHostedService
    {
      
        public Task StartAsync(CancellationToken cancellationToken)
        {
            var excel = new AgregarExcel();
            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            throw new NotImplementedException();
        }

        public void ConfigureServices(IServiceCollection services)
        {
            services.AddHostedService<ProcessData>();
        }
    
    }
}
