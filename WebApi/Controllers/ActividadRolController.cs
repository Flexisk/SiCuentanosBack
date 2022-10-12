using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ActividadRolController : ControllerBase
    {
        private readonly IGenericService<ActividadRol> _service;
        public ActividadRolController(IGenericService<ActividadRol> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<ActividadRol>>> GetActividadRol()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "ok", "No se encontraron ActividadRoles en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<ActividadRol>> GetActividadRol(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron ActividadRol con este Id");
            }

            var ActividadRol = await _service.GetAsync(e => e.RolId == Id, e => e.OrderBy(e => e.RolId), "");

            if (ActividadRol.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron ActividadRol con este Id");
            }
            return Created("ObtenerActividad", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado", ActividadRol = ActividadRol });
        }

        [HttpPost]
        public async Task<ActionResult<ActividadRol>> PostActividadRol(ActividadRol actividadrol)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible Modificar ActividadRol");
            }
            await _service.CreateAsync(actividadrol);
            return Created("ActualizarActividadRol", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo la ActividadRol con Exito" });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutActividadRol(long Id, ActividadRol actividadrol)
        {
            if (Id != actividadrol.RolId)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No fue posible encontrar el Id de la ActividadRol");
            }
            bool updated = await _service.UpdateAsync(Id, actividadrol);

            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Actualizar la ActividadRol");
            }
            return Created("ActualizarActividadRol", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito la ActividadRol" });
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteActividadRol(long Id)
        {
            var ActividadRol = await _service.GetAsync(e => e.RolId == Id, e => e.OrderBy(e => e.RolId), "");

            if (ActividadRol.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron ActividadRol con este Id");
            }
           
            await _service.DeleteAsync(Id);

            return Created("EliminarActividadRol", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito la ActividadRol" });
        }
    }
}
