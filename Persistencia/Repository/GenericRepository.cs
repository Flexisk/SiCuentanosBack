
using Microsoft.EntityFrameworkCore;
using System.Linq.Expressions;
using Persistencia.Context;

namespace Persistencia.Repository
{
    public interface IGenericRepository<T> where T : class
    {
        Task<List<T>> GetAsync();
        
        Task<List<T>> GetAsync(Expression<Func<T, bool>> whereCondition = null,
                           Func<IQueryable<T>, IOrderedQueryable<T>> orderBy = null,
                           string includeProperties = "");
        Task<bool> CreateAsync(T entity);    
        
        Task<bool> UpdateAsync(long id,T entity);

        Task<bool> DeleteAsync(long id);

        bool ExistsAsync();
    }

    public class GenericRepository<T> : IGenericRepository<T> where T : class
    {
        

        private readonly ApplicationDbContext _context;


        public GenericRepository(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<List<T>> GetAsync()
        {
            //return await _context.Actividad.ToListAsync();
            return await _context.Set<T>().ToListAsync();
        }

        public async Task<List<T>> GetAsync(Expression<Func<T, bool>> whereCondition = null,
                                  Func<IQueryable<T>, IOrderedQueryable<T>> orderBy = null,
                                  string includeProperties = "")
        {
            IQueryable<T> query = _context.Set<T>();

            if (whereCondition != null)
            {
                query = query.Where(whereCondition);
            }

            foreach (var includeProperty in includeProperties.Split
                (new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
            {
                query = query.Include(includeProperty);
            }

            if (orderBy != null)
            {
                return await orderBy(query).ToListAsync();
            }
            else
            {
                return await query.ToListAsync();
            }
        }

        public async Task<bool> CreateAsync(T entity)
        {
            bool created = false;

            try
            {
                var save = await _context.Set<T>().AddAsync(entity);

                if (save != null)                                  
                    created = true;  

                await  _context.SaveChangesAsync();              
            }
            catch (Exception)
            {
                throw;
            }                        
            return created;
        }

        public async Task<bool> UpdateAsync(long id,T entity)
        { 
            bool edited = false;
            _context.Entry(entity).State = EntityState.Modified;

            try
            {
                await _context.SaveChangesAsync();
                edited = true;
            }
            catch (Exception)
            {
                throw;
        
            }   
            return edited;         
        }

         public async Task<bool> DeleteAsync(long id)
        { 
            bool deleted = false;
            try
            {
                var entity = await _context.Set<T>().FindAsync(id);
                if (entity == null)
                    return deleted;
                

                var removeEntity = _context.Set<T>().Remove(entity);
                if (removeEntity != null)
                {
                    deleted = true;
                }
                await  _context.SaveChangesAsync();
            } 
            catch (Exception)
            {
                throw;
            }

            return deleted;
        }

        public bool ExistsAsync()
        {
            bool existe = true;
            if (_context.Set<T>() == null)
            {
                existe = false;
            } 
            return existe;        
        }
    }
}