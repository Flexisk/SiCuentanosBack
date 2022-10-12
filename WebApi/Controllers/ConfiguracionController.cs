using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ConfiguracionController : ControllerBase
    {
        private readonly IGenericService<Configuracion> _service;
        public ConfiguracionController(IGenericService<Configuracion> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Configuracion>>> GetConfiguracion()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "ok", "No se encontraron Configuraciones en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<Configuracion>> GetConfiguracion(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Configuraciones con este Id");
            }

            var Configuracion = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (Configuracion.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Configuraciones con este Id");
            }
            return Created("ObtenerConfiguracion", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado", Configuracion = Configuracion });
        }

        [HttpPost]
        public async Task<ActionResult<Configuracion>> PostConfiguracion(Configuracion configuracion)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible Modificar Configuracion");
            }
            await _service.CreateAsync(configuracion);
            return Created("ActualizarConfiguracion", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo la Configuracion con Exito" });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutConfiguracion(long Id, Configuracion configuracion)
        {
            if (Id != configuracion.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No fue posible encontrar el Id de la Configuracion");
            }
            bool updated = await _service.UpdateAsync(Id, configuracion);

            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Actualizar la Configuracion");
            }
            return Created("ActualizarConfiguracion", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito la Configuracion" });
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteConfiguracion(long Id)
        {
            var Configuracion = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (Configuracion.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Configuraciones con este Id");
            }
            await _service.DeleteAsync(Id);

            return Created("EliminarConfiguracion", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito la Configuracion" });
        }
    }
}
