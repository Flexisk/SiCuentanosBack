using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ParametroController : ControllerBase
    {
        private readonly IGenericService<Parametro> _service;
        public ParametroController(IGenericService<Parametro> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Parametro>>> GetParametro()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "ok", "No se encontraron Parametros en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<Parametro>> GetParametro(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Parametros con este Id");
            }

            var Parametro = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (Parametro.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Parametros con este Id");
            }
            return Created("ObtenerParametro", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado", Parametro = Parametro });
        }

        [HttpPost]
        public async Task<ActionResult<Parametro>> PostParametro(Parametro parametro)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible Modificar Parametro");
            }
            await _service.CreateAsync(parametro);

            return Created("ActualizarParametro", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo el Parametro con Exito" });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutParametro(long Id, Parametro parametro)
        {
            if (Id != parametro.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No fue posible encontrar el Id del Parametro");
            }
            bool updated = await _service.UpdateAsync(Id, parametro);

            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Actualizar el Parametro");
            }
            return Created("ActualizarParametro", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito el Parametro" });
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteParametro(long Id)
        {
            var Parametro = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (Parametro.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Parametros con este Id");
            }
            await _service.DeleteAsync(Id);

            return Created("EliminarParametro", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito el Parametro" });
        }
    }
}
