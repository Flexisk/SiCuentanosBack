using Aplicacion.ManejadorErrores;
using Aplicacion.Services;
using Dominio.Administrador;
using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ContratoController : ControllerBase
    {
        private readonly IGenericService<Contrato> _service;
        public ContratoController(IGenericService<Contrato> service)
        {
            this._service = service;
        }

        [HttpGet]
        public async Task<ActionResult<IEnumerable<Contrato>>> GetContrato()
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NoContent, "ok", "No se encontraron Contratos en Base de Datos");
            }
            return await _service.GetAsync();
        }

        [HttpGet("{Id}")]
        public async Task<ActionResult<Contrato>> GetContrato(long Id)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Contratos con este Id");
            }

            var Contrato = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (Contrato.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Contratos con este Id");
            }

            return Created("ObtenerContrato", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se obtuvo Id Solicitado", Contrato = Contrato });
        }

        [HttpPost]
        public async Task<ActionResult<Contrato>> PostContrato(Contrato contrato)
        {
            if (!_service.ExistsAsync())
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Modificado", "No fue posible Modificar Contrato");
            }
            await _service.CreateAsync(contrato);

            return Created("ActualizarContrato", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Creo el Contrato con Exito" });
        }

        [HttpPut("{Id}")]
        public async Task<IActionResult> PutContrato(long Id, Contrato contrato)
        {
            if (Id != contrato.Id)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No fue posible encontrar el Id del Contrato");
            }
            bool updated = await _service.UpdateAsync(Id, contrato);

            if (!updated)
            {
                throw new ExcepcionError(HttpStatusCode.NotModified, "No Modificado", "No se pudo Actualizar el Contrato");
            }
            return Created("ActualizarContrato", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Actualizo con Exito el Contrato" });
        }

        [HttpDelete("{Id}")]
        public async Task<IActionResult> DeleteContrato(long Id)
        {
            var Contrato = await _service.GetAsync(e => e.Id == Id, e => e.OrderBy(e => e.Id), "");

            if (Contrato.Count < 1)
            {
                throw new ExcepcionError(HttpStatusCode.NotFound, "No Existe", "No se encontraron Contratos con este Id");
            }

            await _service.DeleteAsync(Id);

            return Created("EliminarContrato", new { Codigo = HttpStatusCode.OK, Titulo = "OK", Mensaje = "Se Elimino con Exito el Contrato" });
        }
    }
}
