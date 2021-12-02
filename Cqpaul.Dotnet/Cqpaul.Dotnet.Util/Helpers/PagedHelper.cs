using Microsoft.EntityFrameworkCore;

namespace Cqpaul.Dotnet.Util.Helpers
{
    public static class PagedHelper
    {
        /// <summary>
        /// 分页拓展
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <param name="pageIndex">页码，必须大于0</param>
        /// <param name="pageSize"></param>
        /// <returns></returns>
        public static PageResult<TEntity> GetPagedResult<TEntity>(this IQueryable<TEntity> entities, int pageIndex = 1, int pageSize = 20)
        {
            if (pageIndex <= 0) throw new InvalidOperationException($"{nameof(pageIndex)} 必须是大于0的正整数。");

            var totalCount = entities.Count();
            var items = entities.Skip((pageIndex - 1) * pageSize).Take(pageSize).ToList();
            var totalPages = (int)Math.Ceiling(totalCount / (double)pageSize);

            return new PageResult<TEntity>
            {
                PageNo = pageIndex,
                PageSize = pageSize,
                Rows = items,
                TotalRows = totalCount,
                TotalPage = totalPages
            };
        }

        /// <summary>
        /// 分页拓展
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <param name="pageIndex">页码，必须大于0</param>
        /// <param name="pageSize"></param>
        /// <param name="cancellationToken"></param>
        /// <returns></returns>
        public static async Task<PageResult<TEntity>> GetPagedResultAsync<TEntity>(this IQueryable<TEntity> entities, int pageIndex = 1, int pageSize = 20, CancellationToken cancellationToken = default)
        {
            if (pageIndex <= 0) throw new InvalidOperationException($"{nameof(pageIndex)} 必须是大于0的正整数。");

            var totalCount = await entities.CountAsync(cancellationToken);
            var items = await entities.Skip((pageIndex - 1) * pageSize).Take(pageSize).ToListAsync(cancellationToken);
            var totalPages = (int)Math.Ceiling(totalCount / (double)pageSize);

            return new PageResult<TEntity>
            {
                PageNo = pageIndex,
                PageSize = pageSize,
                Rows = items,
                TotalRows = totalCount,
                TotalPage = totalPages
            };
        }

        /// <summary>
        /// 分页拓展
        /// </summary>
        /// <typeparam name="TEntity"></typeparam>
        /// <param name="entities"></param>
        /// <param name="pageIndex">页码，必须大于0</param>
        /// <param name="pageSize"></param>
        /// <returns></returns>
        public static PageResult<TEntity> GetPagedResult<TEntity>(this IEnumerable<TEntity> entities, int pageIndex = 1, int pageSize = 20)
        {
            if (pageIndex <= 0) throw new InvalidOperationException($"{nameof(pageIndex)} 必须是大于0的正整数。");

            var totalCount = entities.Count();
            var items = entities.Skip((pageIndex - 1) * pageSize).Take(pageSize).ToList();
            var totalPages = (int)Math.Ceiling(totalCount / (double)pageSize);

            return new PageResult<TEntity>
            {
                PageNo = pageIndex,
                PageSize = pageSize,
                Rows = items,
                TotalRows = totalCount,
                TotalPage = totalPages
            };
        }
    }

    public class PageResult<T>
    {
        public int PageNo { get; set; }
        public int PageSize { get; set; }
        public int TotalPage { get; set; }
        public int TotalRows { get; set; }
        public ICollection<T>? Rows { get; set; }
    }

}
