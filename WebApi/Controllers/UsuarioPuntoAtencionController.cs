using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UsuarioPuntoAtencionController : ControllerBase
    {

        private readonly IGenericService<UsuarioPuntoAtencion> _service;
        public UsuarioPuntoAtencionController(IGenericService<UsuarioPuntoAtencion> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<UsuarioPuntoAtencion>>> GetUsuarioPuntoAtencion()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "No Existe", "No se encontraron UsuarioPuntoAtencion en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<UsuarioPuntoAtencion>> GetUsuarioPuntoAtencionId(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron UsuarioPuntoAtencion con este Id");
            }
            var usuariopuntoatencion = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), ""); 
            if (usuariopuntoatencion.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron UsuarioPuntoAtencion con este Id");
            }
            return Created("ObtenerUsuarioPuntoAtencion", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado", usuariopuntoatencion = usuariopuntoatencion });
        }

        [HttpPost]
        public async Task<ActionResult<UsuarioPuntoAtencion>> PostUsuarioPuntoAtencion(UsuarioPuntoAtencion usuariopuntoatencion)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible Modificar UsuarioPuntoAtencion");
            }
            await _service.CreateAsync(usuariopuntoatencion);
            return Created("CrearUsuarioPuntoAtencion", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo la UsuarioPuntoAtencion con Exito" });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutUsuarioPuntoAtencion(long Id, UsuarioPuntoAtencion usuariopuntoatencion)
        {
            if (Id != usuariopuntoatencion.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No fue posible encontrar el Id de  UsuarioPuntoAtencion");
            }
            bool updated = await _service.UpdateAsync(Id, usuariopuntoatencion);
            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Actualizar UsuarioPuntoAtencion");
            }
            return Created("ActualizarUsuarioPuntoAtencion", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito la UsuarioPuntoAtencion" });
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteUsuarioPuntoAtencion(long Id)
        {
            var usuariopuntoatencion = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (usuariopuntoatencion.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron UsuarioPuntoAtencion con este Id");
            }
            await _service.DeleteAsync(Id);
            return Created("EliminarUsuarioPuntoAtencion", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito la UsuarioPuntoAtencion" });
        }
    }
}
