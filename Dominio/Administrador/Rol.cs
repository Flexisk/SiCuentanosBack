using System.Numerics;

namespace Dominio.Administrador
{
    public class Rol
    {
        public long Id { get; set; }

        public long ModuloId { get; set; }

        public String? VcNombre { get; set; }

        public Boolean BEstado { get; set; }

        public String? VcCodigoInterno { get; set; }

        public DateTime DtFechaCreacion { get; set; }

        public DateTime? DtFechaActualizacion { get; set; }

        public DateTime? DtFechaAnulacion { get; set; }

        public Modulo? Modulo { get; set; }

        public virtual ICollection<Actividad>? Actividades { get; set; }
        public virtual List <ActividadRol>? ActividadRoles { get; set; }

    }
}
