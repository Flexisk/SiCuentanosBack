using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UsuarioAreaController : ControllerBase
    {
        private readonly IGenericService<UsuarioArea> _service;
        public UsuarioAreaController(IGenericService<UsuarioArea> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<UsuarioArea>>> GetUsuarioArea()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "No Existe", "No se encontro UsuarioArea en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<UsuarioArea>> GetUsuarioAreaId(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontro UsuarioArea con este Id");
            }

            var usuarioarea = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (usuarioarea.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontro UsuarioArea con este Id");
            }
            return Created("ObtenerUsuarioArea", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id UsuarioArea Solicitado",usuarioarea=usuarioarea});
        }

        [HttpPost]
        public async Task<ActionResult<UsuarioArea>> PostUsuarioArea(UsuarioArea usuarioarea)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No fue posible Modificar UsuarioArea");
            }
            await _service.CreateAsync(usuarioarea);
            return Created("CrearUsuarioArea", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo UsuarioArea con Exito" });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutUsuarioArea(long Id, UsuarioArea usuarioarea)
        {
            if (Id != usuarioarea.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "No Existe", "No fue posible encontrar el Id del UsuarioArea");
            }
            bool updated = await _service.UpdateAsync(Id, usuarioarea);
            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Modificar  UsuarioArea");
            }
            return Created("ActualizarUsuarioArea", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito UsuarioArea" });
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteUsuarioArea(long Id)
        {
            var usuarioarea = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (usuarioarea.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontro UsuarioArea con este Id");
            }
            await _service.DeleteAsync(Id);
            return Created("EliminarUsuarioArea", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito  UsuarioArea" });
        }
    }

}

