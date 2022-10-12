using System.Numerics;

namespace Dominio.Administrador
{
    public class ParametroDetalle
    {
        public long Id { get; set; }

        public long ParametroId { get; set; }

        public Parametro? Parametro { get; set; }

        public String? VcNombre { get; set; }
            
        public String? TxDescripcion { get; set; }

        public String? VcCodigoInterno { get; set; }

        public Decimal? DCodigoIterno { get; set; }

        public Boolean BEstado { get; set; }

        public int? RangoDesde { get; set; }

        public int? RangoHasta { get; set; }

        public virtual List <UsuarioCargo>? UsuarioCargos { get; set; }
        public virtual ICollection<Usuario>? CargosUsuario { get; set; }
        public virtual List <UsuarioArea>? UsuarioAreas { get; set; }
        public virtual ICollection<Usuario>? AreasUsuario { get; set; }
        public virtual List <UsuarioPuntoAtencion>? UsuarioPuntoAtenciones { get; set; }
        public virtual ICollection<Usuario>? PuntoAtencionesUsuarios { get; set; }

    }
}
