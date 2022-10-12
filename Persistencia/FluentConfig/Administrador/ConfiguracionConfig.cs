using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using Dominio.Administrador;

namespace Persistencia.FluentConfig.Administrador
{
    public class ConfiguracionConfig
    {
        public ConfiguracionConfig(EntityTypeBuilder<Configuracion> entity)
        {
            entity.ToTable("Configuracion");
            entity.HasKey(p=> p.Id); 

            entity.Property(p=> p.IDiasLimiteRespuesta).IsRequired().HasMaxLength(100);

            entity.Property(p => p.BSabadoLaboral).IsRequired().HasMaxLength(50);

            entity.Property(p=> p.BDomingoLaboral).IsRequired().HasMaxLength(50);

            entity.Property(p=> p.BEstado).IsRequired();

            entity.Property(p=> p.DtFechaCreacion).IsRequired();

            entity.Property(p=> p.DtFechaActualizacion).IsRequired(false);

            entity.Property(p=> p.DtFechaAnulacion).IsRequired(false);
        }
    }
}