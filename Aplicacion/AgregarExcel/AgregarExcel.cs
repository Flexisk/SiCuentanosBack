using Aplicacion.ManejadorErrores;
using Dominio.Administrador;
using Microsoft.EntityFrameworkCore;
using Persistencia.Context;
using SpreadsheetLight;
using System.Data.SqlClient;
using System.Net;
using System.Transactions;

namespace Aplicacion.AgregarExcel
{
    public class AgregarExcel
    { 
        public AgregarExcel()
        {
           parametro();
           //parametrodetalle();
        }

        //public void parametro()
        //{

        //    var connection = new SqlConnection("Data Source=LAPTOP-SOA-46\\SQLEXPRESS;Initial Catalog=ModuloAdministradorSC;Integrated Security=True;");
        //    connection.Open();
        //    SLDocument sL = new SLDocument(@"C:\Users\erojas\Pictures\SICUENTANOS_Back\PARAMETROSSICUENTANOS.xlsx", "PARAMETRO");

        //    using (var transaction = new CommittableTransaction(new TransactionOptions { IsolationLevel = IsolationLevel.ReadCommitted }))

        //    using (var dataContext = new DataModelContainer(builder ))




        //    try
        //    {
        //        var options = new DbContextOptionsBuilder<ApplicationDbContext>().UseSqlServer(connection).Options;

        //        using (var context = new ApplicationDbContext(options))
        //        {
        //            context.Database.UseTransaction(transaction);

        //            context.Database.ExecuteSqlRaw("SET IDENTITY_INSERT [dbo].[Parametro] ON");

        //            int iRow = 2;
        //            while (!string.IsNullOrEmpty(sL.GetCellValueAsString(iRow, 1)))
        //            {
        //                //long Id = sL.GetCellValueAsInt64(iRow, 1);
        //                string VcNombre = sL.GetCellValueAsString(iRow, 2);
        //                string VcCodigoInterno = sL.GetCellValueAsString(iRow, 3);
        //                Boolean BEstado = sL.GetCellValueAsBoolean(iRow, 4);
        //                DateTime DtFechaActualizacion = sL.GetCellValueAsDateTime(iRow, 5);
        //                DateTime DtFechaAnulacion = sL.GetCellValueAsDateTime(iRow, 6);

        //                var parametro = new Parametro();

        //                //parametro.Id = Id;
        //                parametro.VcNombre = VcNombre;
        //                parametro.VcCodigoInterno = VcCodigoInterno;
        //                parametro.BEstado = BEstado;
        //                parametro.DtFechaActualizacion = DtFechaActualizacion;

        //                context.Parametro.Add(parametro);
        //                context.SaveChanges();
        //                iRow++;
        //            }
        //            context.Database.ExecuteSqlRaw("SET IDENTITY_INSERT [dbo].[Parametro] OFF");
        //            transaction.Commit();
        //        }
        //    }

        //    catch (Exception)
        //    {
        //        transaction.Rollback("Parametro");
        //        return;
        //    }
        //}

        //public List<string> parametrodetalle()
        //{

        //    var connection = new SqlConnection("Data Source=LAPTOP-SOA-46\\SQLEXPRESS;Initial Catalog=ModuloAdministradorSC;Integrated Security=True;");
        //    connection.Open();
        //    SLDocument sL = new SLDocument(@"C:\Users\erojas\Pictures\SICUENTANOS_Back\PARAMETROSSICUENTANOS.xlsx", "PARAMETRO_DETALE");


        //    using var transaction = connection.BeginTransaction();


        //    var errores = new List<string>();

        //    try
        //    {
        //        var options = new DbContextOptionsBuilder<ApplicationDbContext>().UseSqlServer(connection).Options;

        //        using (var context = new ApplicationDbContext(options))
        //        {
        //            context.Database.UseTransaction(transaction);

        //            context.Database.ExecuteSqlRaw("SET IDENTITY_INSERT [dbo].[ParametroDetalle] ON");


        //            int iRow = 2;
        //            while (!string.IsNullOrEmpty(sL.GetCellValueAsString(iRow, 1)))
        //            {
        //                long Id = sL.GetCellValueAsInt32(iRow, 1);

        //                if (Id <= 0)
        //            {
        //                errores.Add("El Id del Parametro no puede ser menor a cero en la celda (A" + iRow + ")");
        //            }
        //            long ParametroId = sL.GetCellValueAsInt32(iRow, 2);
        //            string VcNombre = sL.GetCellValueAsString(iRow, 3);
        //            string TxDescripcion = sL.GetCellValueAsString(iRow, 4);
        //            string VcCodigoInterno = sL.GetCellValueAsString(iRow, 6);
        //            decimal DCodigoIterno = sL.GetCellValueAsDecimal(iRow, 7);
        //            Boolean BEstado = sL.GetCellValueAsBoolean(iRow, 8);
        //            int RangoDesde = sL.GetCellValueAsInt32(iRow, 9);
        //            int RangoHasta = sL.GetCellValueAsInt32(iRow, 10);


        //            var paramdetalle = new ParametroDetalle();

        //            paramdetalle.Id = Id;
        //            paramdetalle.ParametroId = ParametroId;
        //            paramdetalle.VcNombre = VcNombre;
        //            paramdetalle.TxDescripcion = TxDescripcion;
        //            paramdetalle.VcCodigoInterno = VcCodigoInterno;
        //            paramdetalle.DCodigoIterno = DCodigoIterno;
        //            paramdetalle.BEstado = BEstado;
        //            paramdetalle.RangoDesde = RangoDesde;
        //            paramdetalle.RangoHasta = RangoHasta;

        //            context.ParametroDetalle.Add(paramdetalle);
        //            context.SaveChanges();
        //            iRow++;
        //        }


        //        context.Database.ExecuteSqlRaw("SET IDENTITY_INSERT [dbo].[Parametro] OFF");

        //            if (errores.Count > 0 )
        //            {
        //                transaction.Rollback("Parametro");
        //            }else
        //            {
        //                transaction.Commit();
        //            }
        //            return errores;

        //        }
        //    }

        //    catch (Exception)
        //    {
        //        transaction.Rollback("Parametro");
        //        return errores;
        //    }
        //}

        public List<string> parametro()
        {
            //using (var transactio = new CommittableTransaction(new TransactionOptions { IsolationLevel = IsolationLevel.ReadCommitted }))

            {
                SLDocument sL = new SLDocument(@"C:\Users\fgarcia\Documents\RepositoriosSiCuentanos\Azure\agregardatos\PARAMETROSNUEVOSICUENTANOS.xlsx", "PARAMETRO");
                using var connection = new SqlConnection("Data Source=LAPTOP-SOA-47;Initial Catalog=ModuloAdministradorSC;Integrated Security=True;");connection.Open();
                using var transactions = connection.BeginTransaction();  
                                
                var errores = new List<string>();

                try
                {
                    var options = new DbContextOptionsBuilder<ApplicationDbContext>()
                   .UseSqlServer(connection)
                   .Options;

                    using (var context = new ApplicationDbContext(options))
                    {
                        context.Database.OpenConnection();
                        context.Database.UseTransaction(transactions);
                        context.Database.ExecuteSqlRaw("SET IDENTITY_INSERT [dbo].[Parametro] ON");

                        int iRow = 2;
                        while (!string.IsNullOrEmpty(sL.GetCellValueAsString(iRow, 1)))
                        {
                            var parametro = new Parametro();

                            long Id = sL.GetCellValueAsInt64(iRow, 1);                       
                            string VcNombre = sL.GetCellValueAsString(iRow, 2);
                            string VcCodigoInterno = sL.GetCellValueAsString(iRow, 3);
                            Boolean BEstado = sL.GetCellValueAsBoolean(iRow, 4);
                            DateTime DtFechaActualizacion = sL.GetCellValueAsDateTime(iRow, 5);
                            DateTime DtFechaAnulacion = sL.GetCellValueAsDateTime(iRow, 6);

                            parametro.Id = Id;
                            if (Id <= 0)
                            {
                                errores.Add("El Id del Parametro no puede ser menor a cero en la celda (A" + iRow + ")");
                                //throw new ExcepcionError(HttpStatusCode.NoContent, "No Existe", "El Id del Parametro no puede ser menor a cero en la celda (A" + iRow +")");
                            }

                            parametro.VcNombre = VcNombre;
                            if (VcNombre == null)
                            {
                                errores.Add("VcNombre no puede ser vacio en la celda (B" + iRow + ")");
                            }

                            parametro.VcCodigoInterno = VcCodigoInterno;
                            if (VcCodigoInterno == null)
                            {
                                errores.Add("VcCodigoInterno no puede ser vacio en la celda (C" + iRow + ")");
                            }

                            parametro.BEstado = BEstado;
                            if (BEstado == null)
                            {
                                errores.Add("BEstado no puede ser vacio en la celda (D" + iRow + ")");
                            }

                            parametro.DtFechaActualizacion = DtFechaActualizacion;

                            context.Parametro.Add(parametro);
                            context.SaveChanges();
                            iRow++;
                        }

                        context.Database.ExecuteSqlRaw("SET IDENTITY_INSERT [dbo].[Parametro] OFF");
                        transactions.Commit();

                        if (errores.Count > 0)
                        {
                            transactions.Rollback();
                            Console.WriteLine(errores);
                        }
                        else
                        {
                            transactions.Commit();
                        }
                        return errores;
                    }
                }
                catch (Exception)
                {
                    return errores;
                }
                finally
                {
                    transactions.Rollback();
                    connection.Close();
                }

            }
        }
        //public List<string> parametrodetalle()
        //{
        //    //using (var transactions = new CommittableTransaction(new TransactionOptions { IsolationLevel = IsolationLevel.ReadCommitted }))

        //    {
        //        SLDocument sL = new SLDocument(@"C:\Users\fgarcia\Documents\RepositoriosSiCuentanos\Azure\agregardatos\PARAMETROSNUEVOSICUENTANOS.xlsx", "PARAMETRO_DETALLE");
        //        using var connection = new SqlConnection("Data Source=LAPTOP-SOA-47;Initial Catalog=ModuloAdministradorSC;Integrated Security=True;"); connection.Open();
        //        using var transaction = connection.BeginTransaction();

        //        var errores = new List<string>();

        //        try
        //        {
        //            var options = new DbContextOptionsBuilder<ApplicationDbContext>()
        //           .UseSqlServer(connection)
        //           .Options;

        //            using (var context = new ApplicationDbContext(options))
        //            {

        //                context.Database.OpenConnection();
        //                context.Database.UseTransaction(transaction);
        //                context.Database.ExecuteSqlRaw("SET IDENTITY_INSERT [dbo].[ParametroDetalle] ON");

        //                int iRow = 2;
        //                while (!string.IsNullOrEmpty(sL.GetCellValueAsString(iRow, 1)))
        //                {
        //                    long Id = sL.GetCellValueAsInt64(iRow, 1);
        //                    long ParametroId = sL.GetCellValueAsInt32(iRow, 2);
        //                    string VcNombre = sL.GetCellValueAsString(iRow, 3);
        //                    string TxDescripcion = sL.GetCellValueAsString(iRow, 4);
        //                    string VcCodigoInterno = sL.GetCellValueAsString(iRow, 6);
        //                    decimal DCodigoIterno = sL.GetCellValueAsDecimal(iRow, 7);
        //                    Boolean BEstado = sL.GetCellValueAsBoolean(iRow, 8);
        //                    int RangoDesde = sL.GetCellValueAsInt32(iRow, 9);
        //                    int RangoHasta = sL.GetCellValueAsInt32(iRow, 10);

        //                    var parametrodetalles = new ParametroDetalle();

        //                    parametrodetalles.Id = Id;
        //                    if (Id <= 0)
        //                    {
        //                        errores.Add("El Id del ParametroDetalle no puede ser menor a cero en la celda (A" + iRow + ")");
        //                        throw new ExcepcionError(HttpStatusCode.NoContent, "No Existe", "El Id del Parametro no puede ser menor a cero en la celda (A" + iRow + ")");
        //                    }
        //                    parametrodetalles.ParametroId = ParametroId;
        //                    if (ParametroId == null)
        //                    {
        //                        errores.Add("ParametroId no puede ser vacio en la celda (B" + iRow + ")");
        //                    }
        //                    parametrodetalles.VcNombre = VcNombre;
        //                    if (VcNombre == null)
        //                    {
        //                        errores.Add("VcNombre no puede ser vacio en la celda (C" + iRow + ")");
        //                    }
        //                    parametrodetalles.TxDescripcion = TxDescripcion;
        //                    if (TxDescripcion == null)
        //                    {
        //                        errores.Add("TxDescripcion no puede ser vacio en la celda (D" + iRow + ")");
        //                    }
        //                    parametrodetalles.VcCodigoInterno = VcCodigoInterno;
        //                    parametrodetalles.DCodigoIterno = DCodigoIterno;
        //                    parametrodetalles.BEstado = BEstado;
        //                    parametrodetalles.RangoDesde = RangoDesde;
        //                    parametrodetalles.RangoHasta = RangoHasta;

        //                    context.ParametroDetalle.Add(parametrodetalles);
        //                    context.SaveChanges();
        //                    iRow++;
        //                }

        //                context.Database.ExecuteSqlRaw("SET IDENTITY_INSERT [dbo].[ParametroDetalle] OFF");
        //                transaction.Commit();

        //                if (errores.Count > 0)
        //                {
        //                    transaction.Rollback();
        //                    Console.WriteLine(errores);

        //                }
        //                else
        //                {
        //                    transaction.Commit();
        //                }
        //                return errores;
        //            }
        //        }
        //        catch (Exception)
        //        {
                    
        //            return errores;
        //        }
        //        finally
        //        {
        //            //transaction.Rollback();
        //            connection.Close();
        //        }
        //    }
        //}
    }
}






