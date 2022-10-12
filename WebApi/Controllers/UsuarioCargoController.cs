using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UsuarioCargoController : ControllerBase
    {
        private readonly IGenericService<UsuarioCargo> _service;
        public UsuarioCargoController(IGenericService<UsuarioCargo> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<UsuarioCargo>>> GetUsuarioCargo()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "No Existe", "No se encontraron UsuarioCargo en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<UsuarioCargo>> GetUsuarioCargoId(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron UsuarioCargo con este Id");
            }

            var usuariocargo = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (usuariocargo.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron UsuarioCargo con este Id");
            }
            return Created("ObtenerUsuarioCargo", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado", usuariocargo = usuariocargo });
        }

        [HttpPost]
        public async Task<ActionResult<UsuarioCargo>> PostUsuarioCargo(UsuarioCargo usuariocargo)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible  Crear UsuarioCargo");
            }
            await _service.CreateAsync(usuariocargo);
            return Created("CrearUsuarioCargo", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo la UsuarioCargo con Exito"});
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutUsuarioCargo(long Id, UsuarioCargo usuariocargo)
        {
            if (Id != usuariocargo.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No fue posible encontrar Id de UsuarioCargo");
            }
            bool updated = await _service.UpdateAsync(Id, usuariocargo);
            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Actualizar  UsuarioCargo");
            }
            return Created("ActualizarUsuarioCargo", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito la UsuarioCargo" });
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteUsuarioCargo(long Id)
        {

            var usuariocargo = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (usuariocargo.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontro UsuarioCargo con este Id");
            }
            await _service.DeleteAsync(Id);
            return Created("EliminarUsuariocargo", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito UsuarioCargo" });
        }









    }
}
