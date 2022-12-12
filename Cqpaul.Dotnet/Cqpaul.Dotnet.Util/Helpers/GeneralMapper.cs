using Cqpaul.Dotnet.Util.Extensions;
using Cqpaul.Dotnet.Util.Model;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cqpaul.Dotnet.Util.Helpers
{
    public static class GeneralMapper
    {
        /// <summary>
        /// 将数据Rows，转换成为对应的对象List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <param name="isOnlyUserInputInfo">是否仅仅取用户上传的数据，进行赋值</param>
        /// <returns></returns>
        public static List<T> ExcelSheetRowsDataToModelList<T>(List<IRow> rows, bool isOnlyUserInputInfo = true) where T : BaseSheetModel, new()
        {
            List<T> dataModels = new List<T>();
            foreach (IRow row in rows)
            {
                T tableItem = new T();
                tableItem.SetObjectValueByUploadDataRow(row, isOnlyUserInputInfo);
                dataModels.Add(tableItem);
            }
            return dataModels;
        }

        /// <summary>
        /// 将数据Rows，转换成为对应的对象List
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="rows"></param>
        /// <param name="propertyNames">显示展示Mapping的属性</param>
        /// <returns></returns>
        public static List<T> ExcelSheetRowsDataToModelList<T>(List<IRow> rows, List<string> propertyNames) where T : BaseSheetModel, new()
        {
            List<T> dataModels = new List<T>();
            foreach (IRow row in rows)
            {
                T tableItem = new T();
                tableItem.SetObjectValueByUploadDataRow(row, propertyNames);
                dataModels.Add(tableItem);
            }
            return dataModels;
        }

    }
}
