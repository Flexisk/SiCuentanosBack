using Dominio.Administrador;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Persistencia.FluentConfig.Administrador
{
    public class UsuarioConfig
    {
        public UsuarioConfig(EntityTypeBuilder<Usuario> entity)
        {

            entity.ToTable("Usuario");
            entity.HasKey(p => p.Id);

            entity
           .HasMany(p => p.Contrato)
           .WithOne(p => p.Usuario)
           .HasForeignKey(p => p.UsuarioId)
           .HasConstraintName("FK_Usuario_Contrato")
           .OnDelete(DeleteBehavior.Restrict);

            entity.Property(p => p.TipoDocumento).IsRequired().HasMaxLength(50);

            entity.Property(p => p.VcDocumento).IsRequired().HasMaxLength(80);

            entity.Property(p => p.VcPrimerNombre).IsRequired().HasMaxLength(100);

            entity.Property(p => p.VcPrimerApellido).IsRequired().HasMaxLength(100);

            entity.Property(p => p.VcSegundoNombre)
                .IsRequired(false)
                .HasMaxLength(100)
                .HasComment("Segundo Nombre del usuario");

            entity.Property(p => p.VcSegundoApellido)
                .IsRequired(false)
                .HasMaxLength(100)
                .HasComment("Segundo Apellido del usuario");

            entity.Property(p => p.VcCorreo).IsRequired().HasMaxLength(100);

            entity.Property(p => p.VcDireccion)
                .IsRequired(false)
                .HasMaxLength(100)
                .HasComment("Segundo Apellido del usuario");

            entity.Property(p => p.VcIdAzure).IsRequired().HasMaxLength(100);

            entity.Property(p => p.BEstado).IsRequired().HasMaxLength(100);

            entity.Property(p => p.DtFechaCreacion).IsRequired().HasMaxLength(100);

            entity.Property(p => p.DtFechaActualizacion).IsRequired(false).HasMaxLength(100);

            entity.Property(p => p.DtFechaAnulacion).IsRequired(false).HasMaxLength(100);

        }
    }
}
