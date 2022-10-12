using System.Numerics;

namespace Dominio.Administrador
{
    public class UsuarioCargo
    {
        public long Id { get; set; }

        public long UsuarioId { get; set; }
        public Usuario? Usuario { get; set; }

        public long ParametroDetalleId { get; set; }
        public ParametroDetalle? ParametroDetalle { get; set; }

        public DateTime DtFechaInicio { get; set; }

        public DateTime DtFechaFin { get; set; }

        public DateTime DtFechaCreacion { get; set; }

        public DateTime? DtFechaAnulacion { get; set; }

    }
}
