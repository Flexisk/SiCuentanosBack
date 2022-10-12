using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ActividadController : ControllerBase
    {
        private readonly IGenericService<Actividad> _service;
        public ActividadController(IGenericService<Actividad> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Actividad>>> GetActividad()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "No Existe", "No se encontraron Actividades en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<Actividad>> GetActividad(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Activivades con este Id");
            }
            var Actividad = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");
            if (Actividad.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Activivades con este Id");
            }
            return Created("ObtenerActividad", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado", Actividad = Actividad});
        }

        [HttpPost]
        public async Task<ActionResult<Actividad>> PostActividad(Actividad actividad)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible Modificar Activivad");
            }
            await _service.CreateAsync(actividad);

            return Created("CrearActividad",  new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo la Activivad con Exito", actividad = actividad });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutActividad(long Id, Actividad actividad)
        {
            if (Id != actividad.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No fue posible encontrar el Id de la Activivad");
            }
            bool updated = await _service.UpdateAsync(Id, actividad);

            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Actualizar la Activivad");
            }
            return Created("ActualizarActividad", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito la Activivad"});
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteActividad(long Id)
        {
            var actividad = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (actividad.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Activivades con este Id");
            }
            await _service.DeleteAsync(Id);
            return Created("EliminarActividad", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito la Actividad" });
        }
    }

}
