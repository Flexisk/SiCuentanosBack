using System.Linq.Expressions;
using Persistencia.Repository;
namespace Aplicacion.Services
{
    public interface IGenericService<T> where T : class
    {
         Task<List<T>> GetAsync();

        Task<List<T>> GetAsync(Expression<Func<T, bool>> whereCondition = null,
                           Func<IQueryable<T>, IOrderedQueryable<T>> orderBy = null,
                           string includeProperties = "");

        Task<bool> CreateAsync(T entity);

        Task<bool> UpdateAsync(long id, T entity);

        Task<bool> DeleteAsync(long id);

        bool ExistsAsync();
    }

    public class GenericService<T> : IGenericService<T> where T : class
    {
        public IGenericRepository<T> _genericRepository { get; }

        public GenericService(IGenericRepository<T> genericRepository)
        {
            _genericRepository = genericRepository;
        }
        public async Task<bool> CreateAsync(T entity)
        {
            return await _genericRepository.CreateAsync(entity);
        }


        public async Task<bool> DeleteAsync(long id)
        {
           return await _genericRepository.DeleteAsync(id);
        }

        public async Task<List<T>> GetAsync()
        {
            return await _genericRepository.GetAsync();
        }

        public async Task<List<T>> GetAsync(Expression<Func<T, bool>> whereCondition = null,
                                  Func<IQueryable<T>, IOrderedQueryable<T>> orderBy = null,
                                  string includeProperties = "")
        

      

        {
            return await _genericRepository.GetAsync(whereCondition,orderBy,includeProperties);
        }           

        public async Task<bool> UpdateAsync(long id,T entity)
        {
            return await _genericRepository.UpdateAsync(id,entity);
        }

        public bool ExistsAsync()
        {
            return _genericRepository.ExistsAsync();
        }
    }    

}