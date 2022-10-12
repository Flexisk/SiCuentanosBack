using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class RolController : ControllerBase
    {
        private readonly IGenericService<Rol> _service;
        public RolController(IGenericService<Rol> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Rol>>> GetRol()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "OK", "No se encontro Rol en Base de Datos"); ;
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<Rol>> GetRol(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontro Rol con este Id");
            }

            var Rol = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (Rol.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontro Rol con este Id");
            }
            return Created("ObtenerRol", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado", Rol = Rol });
        }

        [HttpPost]
        public async Task<ActionResult<Rol>> PostRol(Rol rol)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible Modificar Rol");
            }
            await _service.CreateAsync(rol);
            return Created("CrearRol", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo la Rol con Exito" });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutRol(long Id, Rol rol)
        {
            if (Id != rol.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No fue posible encontrar el Id de Rol");
            }
            bool updated = await _service.UpdateAsync(Id, rol);
            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Modificar el Rol");
            }
            return Created("ActualizarRol", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito la Activivad" });
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteRol(long Id)
        {
            var rol = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (rol.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontro Rol con este Id");
            }
            await _service.DeleteAsync(Id);
            return Created("EliminarRol", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino Rol con Exito" });
        }
    }
}
