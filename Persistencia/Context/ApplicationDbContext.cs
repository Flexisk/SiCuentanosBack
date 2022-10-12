using Microsoft.EntityFrameworkCore;
using Dominio.Administrador;
using Persistencia.FluentConfig.Administrador;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Data.Common;

namespace Persistencia.Context
{
    public class ApplicationDbContext : DbContext
    {

        //private DbConnection _connection;


        //public ApplicationDbContext(DbConnection connection)
        //{
        //    this._connection = connection;
        //}

        //public ApplicationDbContext(DbContextOptions options) : base(options)
        //{
        //}

        //protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        //{
        //    optionsBuilder.UseSqlServer("DefaultConnection");
        //}




        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {

        }




        public ApplicationDbContext(string connection) : base(GetOptions(connection, null)) { }

        private static DbContextOptions GetOptions(string connectionString, Action<SqlServerDbContextOptionsBuilder>? sqlServerOptionsAction)
        {
            return SqlServerDbContextOptionsExtensions.UseSqlServer(new DbContextOptionsBuilder(), connectionString, sqlServerOptionsAction).Options;
        }


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            new ModuloConfig(modelBuilder.Entity<Modulo>());
            new ActividadConfig(modelBuilder.Entity <Actividad>());
            new ConfiguracionConfig(modelBuilder.Entity<Configuracion>());
            new ParametroConfig(modelBuilder.Entity<Parametro>());
            new ParametroDetalleConfig(modelBuilder.Entity<ParametroDetalle>());
           
        }

        public DbSet<Actividad> Actividad { get; set; }
        public DbSet<Configuracion> Configuracion { get; set; }
        public DbSet<Contrato> Contrato { get; set; }
        public DbSet<Modulo> Modulo { get; set; }
        public DbSet<Parametro> Parametro { get; set; }
        public DbSet<ParametroDetalle> ParametroDetalle { get; set; }
        public DbSet<Rol> Rol { get; set; }
        public DbSet<Usuario> Usuario { get; set; }
        public DbSet<ActividadRol> ActividadRol { get; set; }
        public DbSet<UsuarioCargo> UsuarioCargo { get; set; }
        public DbSet<UsuarioArea> UsuarioArea { get; set; }
        public DbSet<UsuarioPuntoAtencion> UsuarioPuntoAtencion { get; set; }
    }
}