using System.Numerics;

namespace Dominio.Administrador
{
    public class ActividadRol
    {
        public long ActividadId { get; set; }

        public Actividad? Actividad { get; set; }

        public long RolId { get; set; }

        public Rol? Rol { get; set; }   

    }
}
