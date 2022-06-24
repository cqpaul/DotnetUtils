using System.Reflection;
using Cqpaul.Dotnet.Util.Web.UnitOfWork.Attributes;
using Microsoft.AspNetCore.Mvc.Controllers;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.EntityFrameworkCore;

namespace Cqpaul.Dotnet.Util.Web.UnitOfWork.Filters
{
    public sealed class UnitOfWorkFilter : IAsyncActionFilter
    {
        /// <summary>
        /// 数据库上下文池
        /// </summary>
        private readonly DbContext? _dbContext;

        public async Task OnActionExecutionAsync(ActionExecutingContext context, ActionExecutionDelegate next)
        {
            // 获取动作方法描述器
            var actionDescriptor = context.ActionDescriptor as ControllerActionDescriptor;
            var method = actionDescriptor?.MethodInfo;

            // 判断是否贴有工作单元特性
            if (method == null || !method.IsDefined(typeof(UnitOfWorkAttribute), true))
            {
                // 调用方法
                var resultContext = await next();
            }
            else
            {
                // 获取工作单元特性
                var unitOfWorkAttribute = method.GetCustomAttribute<UnitOfWorkAttribute>();

                if (unitOfWorkAttribute != null && unitOfWorkAttribute.EnsureTransaction)
                {
                    if (_dbContext == null)
                    {
                        throw new Exception($"{nameof(DbContext)} is null.");
                    }
                    using (var tran = _dbContext.Database.BeginTransaction())
                    {
                        try
                        {
                            // 调用方法
                            var resultContext = await next();

                            await tran.CommitAsync();
                        }
                        catch (Exception ex)
                        {
                            await tran.RollbackAsync();
                            throw new Exception(ex.Message);
                        }
                    }
                }
            }

        }
    }
}
