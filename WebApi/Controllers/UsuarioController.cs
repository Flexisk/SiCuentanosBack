using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UsuarioController : ControllerBase
    {
        private readonly IGenericService<Usuario> _service;
        public UsuarioController(IGenericService<Usuario> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Usuario>>> GetUsuario()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "No Existe", "No se encontraron Usuario en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<Usuario>> GetUsuarioId(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Usuarios con este Id");
            }
            var usuario = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");
            if (usuario.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Usuarios con este Id");
            }
            return Created("ObtenerUsuario", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado",usuario=usuario});
        }

        [HttpPost]
        public async Task<ActionResult<Usuario>> PostUsuario(Usuario usuario)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible Modificar Usuario");
            }
            await _service.CreateAsync(usuario);
            return Created("CrearUsuario", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo la Usuario con Exito" });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutUsuario(long Id, Usuario usuario)
        {
            if (Id != usuario.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No fue posible encontrar el Id de Usuario");
            }
            bool updated = await _service.UpdateAsync(Id, usuario);
            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Actualizar  Usuario");
            }
            return Created("ActualizarUsuario", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito Usuario" });
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteUsuario(long Id)
        {
            var usuario = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (usuario.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Usuarios con este Id");
            }
            await _service.DeleteAsync(Id);
            return Created("EliminarUsuario", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito Usuario" });
        }
    }







}
