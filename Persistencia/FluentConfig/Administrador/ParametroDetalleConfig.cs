using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using Dominio.Administrador;

namespace Persistencia.FluentConfig.Administrador
{
    public class ParametroDetalleConfig
    {
        public ParametroDetalleConfig(EntityTypeBuilder<ParametroDetalle> entity)
        {
            entity.ToTable("ParametroDetalle");
            entity.HasKey(p => p.Id);
                   

            entity.HasOne(p => p.Parametro)
                  .WithMany(p => p.ParametroDetalles)

                  .HasForeignKey(p => p.ParametroId)
                  .HasConstraintName("FK_Parametro")
                  .OnDelete(DeleteBehavior.Restrict);


            entity
            .HasMany(p => p.CargosUsuario)
            .WithMany(p => p.Cargos)
            .UsingEntity<UsuarioCargo>(
               j => j
                .HasOne(pt => pt.Usuario)
                .WithMany(t => t.UsuarioCargos)
                .HasForeignKey(f => f.UsuarioId),
               j => j
                .HasOne(pt => pt.ParametroDetalle)
                .WithMany(t => t.UsuarioCargos)
                .HasForeignKey(f => f.ParametroDetalleId),
               j => j
                .HasKey(pt => new { pt.UsuarioId, pt.ParametroDetalleId })
            );

               entity
              .HasMany(p => p.AreasUsuario)
              .WithMany(p => p.Areas)
              .UsingEntity<UsuarioArea>(
                 j => j
                  .HasOne(pt => pt.Usuario)
                  .WithMany(t => t.UsuarioAreas)
                  .HasForeignKey(f => f.UsuarioId),
                 j => j
                  .HasOne(pt => pt.ParametroDetalle)
                  .WithMany(t => t.UsuarioAreas)
                  .HasForeignKey(f => f.ParametroDetalleId),
                 j => j
                  .HasKey(pt => new { pt.UsuarioId, pt.ParametroDetalleId })
              );

                entity
               .HasMany(p => p.PuntoAtencionesUsuarios)
               .WithMany(p => p.PuntoAtenciones)
               .UsingEntity<UsuarioPuntoAtencion>(
                  j => j
                   .HasOne(pt => pt.Usuario)
                   .WithMany(t => t.UsuarioPuntoAtenciones)
                   .HasForeignKey(f => f.UsuarioId),
                  j => j
                   .HasOne(pt => pt.ParametroDetalle)
                   .WithMany(t => t.UsuarioPuntoAtenciones)
                   .HasForeignKey(f => f.ParametroDetalleId),
                  j => j
                   .HasKey(pt => new { pt.UsuarioId, pt.ParametroDetalleId })
               );



            entity.Property(p=> p.VcNombre).IsRequired().HasMaxLength(50);

            entity.Property(p=> p.TxDescripcion).IsRequired(false);

            entity.Property(p=> p.VcCodigoInterno).IsRequired(false).HasMaxLength(50);

            entity.Property(p=> p.DCodigoIterno).IsRequired(false).HasPrecision(17,3);

            entity.Property(p=> p.BEstado).IsRequired();

            entity.Property(p=> p.RangoDesde).IsRequired(false);

            entity.Property(p=> p.RangoHasta).IsRequired(false);
        }
    }
}