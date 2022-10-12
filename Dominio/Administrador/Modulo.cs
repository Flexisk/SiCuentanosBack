using System.Numerics;

namespace Dominio.Administrador
{
    public class Modulo
    {
        public long Id { get; set; }

        public String? VcNombre { get; set; }
        
        public String? VcDescripcion { get; set; }

        public String? VcRedireccion { get; set; }

        public String? VcIcono { get; set; }

        public Boolean BEstado { get; set; }

        public DateTime DtFechaCreacion { get; set; }

        public DateTime? DtFechaActualizacion { get; set; }

        public DateTime? DtFechaAnulacion { get; set; }

        public virtual ICollection<Actividad>? Actividades { get; set; }
        public virtual ICollection<Rol>? Roles { get; set; }
    }
}
