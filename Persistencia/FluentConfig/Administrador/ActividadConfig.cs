using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using Dominio.Administrador;
using System.Numerics;

namespace Persistencia.FluentConfig.Administrador
{
    public class ActividadConfig

    {
        public ActividadConfig(EntityTypeBuilder<Actividad> entity)
        {
            entity.ToTable("Actividad");

            entity.HasKey(p => p.Id);

            entity
            .HasOne(p => p.Modulo)
            .WithMany(p => p.Actividades)

            .HasForeignKey(p => p.ModuloId)
            .HasConstraintName("FK_Actividad_Modulo")
            .OnDelete(DeleteBehavior.Restrict);

            entity
            .HasMany(p => p.Roles)
            .WithMany(p => p.Actividades)
            .UsingEntity<ActividadRol>(
                j => j
                    .HasOne(pt => pt.Rol)
                    .WithMany(t => t.ActividadRoles)
                    .HasForeignKey(f => f.RolId),
                j => j
                    .HasOne(pt => pt.Actividad)
                    .WithMany(t => t.ActividadRoles)
                    .HasForeignKey(f => f.ActividadId),
                j => j
                    .HasKey(pt => new {pt.RolId,pt.ActividadId})
                );

            entity.Property(p => p.VcNombre).IsRequired().HasMaxLength(50);

            entity.Property(p => p.VcDescripcion).IsRequired().HasMaxLength(200);

            entity.Property(p => p.VcRedireccion).IsRequired().HasMaxLength(80);

            entity.Property(p => p.VcIcono).IsRequired().HasMaxLength(20);

            entity.Property(p => p.BEstado).IsRequired();

            entity.Property(p => p.PadreId)
            .IsRequired(false)
            .HasMaxLength(40)
            .HasComment("Id de la actividad padre de acuerdo con la jerarquia");

            entity.Property(p => p.DtFechaCreacion).IsRequired();

            entity.Property(p => p.DtFechaActualizacion).IsRequired(false);

            entity
            .Property(p => p.DtFechaAnulacion)
            .IsRequired(false)
            .HasComment("Fecha anulación del registro");



            entity.HasData(
                new Actividad
                {
                    Id = 1,
                    ModuloId = 1,
                    VcNombre = "Personas",
                    VcDescripcion = "Gestión de personas",
                    VcRedireccion = "#",
                    VcIcono = "bi bi-person",
                    BEstado = true,
                    DtFechaCreacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    DtFechaActualizacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773)
                },
                new Actividad
                {
                    Id = 2,
                    ModuloId = 1,
                    VcNombre = "Roles",
                    VcDescripcion = "Gestión de roles",
                    VcRedireccion = "/actividad",
                    VcIcono = "bi bi-person-rolodex",
                    BEstado = true,
                    DtFechaCreacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    DtFechaActualizacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773)
                },
                new Actividad
                {
                    Id = 3,
                    ModuloId = 1,
                    VcNombre = "Configuración",
                    VcDescripcion = "Configuración General",
                    VcRedireccion = "#",
                    VcIcono = "bi bi-person-rolodex",
                    BEstado = true,
                    DtFechaCreacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    DtFechaActualizacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773)
                },
                new Actividad
                {
                    Id = 4,
                    ModuloId = 1,
                    VcNombre = "Uusarios",
                    VcDescripcion = "Gestión de usuarios",
                    VcRedireccion = "/usuario",
                    VcIcono = "",
                    BEstado = true,
                    DtFechaCreacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    DtFechaActualizacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    

                },
                new Actividad
                {
                    Id = 5,
                    ModuloId = 1,
                    VcNombre = "Cargos",
                    VcDescripcion = "Gestión de Cargos",
                    VcRedireccion = "/cargos",
                    VcIcono = "",
                    BEstado = true,
                    DtFechaCreacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    DtFechaActualizacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    
                },
                new Actividad
                {
                    Id = 6,
                    ModuloId = 1,
                    VcNombre = "Áreas",
                    VcDescripcion = "Gestión de Áreas",
                    VcRedireccion = "#",
                    VcIcono = "",
                    BEstado = true,
                    DtFechaCreacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    DtFechaActualizacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    
                },
                new Actividad
                {
                    Id = 7,
                    ModuloId = 1,
                    VcNombre = "Puntos de Atención",
                    VcDescripcion = "Gestión de Puntos de Atención",
                    VcRedireccion = "#",
                    VcIcono = "bi bi-person",
                    BEstado = true,
                    DtFechaCreacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    DtFechaActualizacion = new DateTime(2022, 8, 13, 11, 15, 9, 749, DateTimeKind.Local).AddTicks(9773),
                    
                }
            );

        }
    }
}