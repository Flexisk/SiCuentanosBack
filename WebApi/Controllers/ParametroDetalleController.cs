using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ParametroDetalleController : ControllerBase
    {
        private readonly IGenericService<ParametroDetalle> _service;
        public ParametroDetalleController(IGenericService<ParametroDetalle> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<ParametroDetalle>>> GetParametroDetalle()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "ok", "No se encontraron ParametroDetalle en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<ParametroDetalle>> GetParametroDetalle(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron ParametroDetalle con este Id");
            }

            var parametrodetalle = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (parametrodetalle.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron ParametroDetalle con este Id");
            }
            return Created("ObtenerParametroDetalle", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado", parametrodetalle = parametrodetalle });
        }

        [HttpPost]
        public async Task<ActionResult<ParametroDetalle>> PostParametroDetalle(ParametroDetalle parametroDetalle)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No fue posible Modificar ParametroDetalle");
            }
            await _service.CreateAsync(parametroDetalle);
            return Created("ActualizarParametroDetalle", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo ParametroDetalle con Exito", parametroDetalle });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutParametroDetalle(long Id, ParametroDetalle parametroDetalle)
        {
            if (Id != parametroDetalle.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible encontrar el Id de ParametroDetalle");
            }
            bool updated = await _service.UpdateAsync(Id, parametroDetalle);
            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Modificar ParametroDetalle");
            }
            return Created("ActualizarActividad", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito ParametroDetalle",parametroDetalle});
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteParametroDetalle(long Id)
        {
            var parametroDetalle = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (parametroDetalle.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No encontraron ParametroDetalle con este Id");
            }
            await _service.DeleteAsync(Id);
            return Created("EliminarParametroDetalle", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito ParametroDetalle" });
        }


    }
}
