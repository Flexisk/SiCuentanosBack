using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;
using System.Reflection;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ModuloController : ControllerBase
    {
        private readonly IGenericService<Modulo> _service;
        public ModuloController(IGenericService<Modulo> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Modulo>>> GetModulo()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "ok", "No se encontraron Modulos en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<Modulo>> GetModulo(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Modulos con este Id");
            }

            var Modulo = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (Modulo.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Modulos con este Id");
            }
            return Created("ObtenerModulo", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado", Modulo = Modulo });
        }

        [HttpPost]
        public async Task<ActionResult<Modulo>> PostModulo(Modulo modulo)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible Modificar Modulo");
            }
            await _service.CreateAsync(modulo);

            return Created("ActualizarModulo", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo la Modulo con Exito" });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutModulo(long Id, Modulo modulo)
        {
            if (Id != modulo.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No fue posible encontrar el Id del Modulo");
            }
            bool updated = await _service.UpdateAsync(Id, modulo);

            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Actualizar el Modulo");
            }
            return Created("ActualizarModulo", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito el Modulo" });
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteModulo(long Id)
        {
            var Modulo = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (Modulo.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Modulos con este Id");
            }
           await _service.DeleteAsync(Id);

            return Created("EliminarModulo", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito el Modulo" });
        }
    }
}
