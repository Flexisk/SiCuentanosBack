using System.Numerics;

namespace Dominio.Administrador
{
    public class Usuario
    {
        public long Id { get; set; }

        public String? TipoDocumento { get; set; }

        public String? VcDocumento { get; set; }

        public String? VcPrimerNombre { get; set; }

        public String? VcPrimerApellido { get; set; }

        public String? VcSegundoNombre { get; set; }

        public String? VcSegundoApellido { get; set; }

        public String? VcCorreo { get; set; }

        public String? VcTelefono { get; set; }

        public String? VcDireccion { get; set; }

        public String? VcIdAzure { get; set; }

        public Boolean BEstado { get; set; }

        public DateTime DtFechaCreacion { get; set; }

        public DateTime? DtFechaActualizacion { get; set; }

        public DateTime? DtFechaAnulacion { get; set; }

        public virtual ICollection <Contrato>? Contrato { get; set; }
        public virtual ICollection<ParametroDetalle>? Cargos { get; set; }
        public virtual List <UsuarioCargo>? UsuarioCargos { get; set; }
        public virtual ICollection<ParametroDetalle>? Areas { get; set; }
        public virtual List <UsuarioArea>? UsuarioAreas { get; set; }
        public virtual ICollection<ParametroDetalle>? PuntoAtenciones { get; set; }
        public virtual List <UsuarioPuntoAtencion>? UsuarioPuntoAtenciones { get; set; }

    }
}
