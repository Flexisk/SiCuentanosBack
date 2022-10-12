using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using Dominio.Administrador;

namespace Persistencia.FluentConfig.Administrador
{
    public class ModuloConfig
    {
        public ModuloConfig(EntityTypeBuilder<Modulo> entity)
        {
            entity.ToTable("Modulo");
            entity.HasKey(p => p.Id);

            entity.HasMany(p=> p.Actividades)
                  .WithOne(p=>p.Modulo);

            entity.HasMany(p => p.Roles)
                  .WithOne(p=>p.Modulo);

        
            entity.Property(p=> p.VcNombre).IsRequired().HasMaxLength(50);

            entity.Property(p=> p.VcDescripcion).IsRequired().HasMaxLength(200);

            entity.Property(p=> p.VcRedireccion).IsRequired(false).HasMaxLength(80);

             entity.Property(p=> p.VcIcono).IsRequired(false).HasMaxLength(50);

            entity.Property(p=> p.BEstado).IsRequired();

            entity.Property(p=> p.DtFechaCreacion).IsRequired();

            entity.Property(p=> p.DtFechaActualizacion).IsRequired();

            entity
            .Property(p=> p.DtFechaAnulacion)
            .IsRequired(false)
            .HasComment("Fecha Eliminacion del registro");



            entity.HasData(
               new Modulo
               {
                   Id = 1,
                   VcNombre = "Administrador",
                   VcDescripcion = "Modulo para gestionar los permisos de los usuarios dentro del aplicativo",
                   VcIcono = "bi bi-sliders",
                   BEstado = true,
                   DtFechaCreacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                   DtFechaActualizacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773)
               },
               new Modulo
               {
                   Id = 2,
                   VcNombre = "Orientación e Información",
                   VcDescripcion = "Modulo para registrar la gestión de orientación e información de la Dirección de Servicio a la Ciudadanía DSC",
                   VcIcono = "bi bi-info-circle-fill",
                   BEstado = true,
                   DtFechaCreacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                   DtFechaActualizacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773)
               },
               new Modulo
               {
                   Id = 3,
                   VcNombre = "Asistencia Técnica",
                   VcDescripcion = "Modulo para realizar seguimiento a las actividades que cumple el equipo de asistencia técnica tales como planes de acción",
                   VcIcono = "bi bi-search-heart",
                   BEstado = true,
                   DtFechaCreacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                   DtFechaActualizacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773)
               }
           );
        }
    }
}