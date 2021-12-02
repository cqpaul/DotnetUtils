using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cqpaul.Dotnet.Util.Extensions
{
    public static class DictionaryExtension
    {
        /// <summary>
        /// 将一个字典转化为 QueryString
        /// </summary>
        /// <param name="dict"></param>
        /// <param name="urlEncode"></param>
        /// <returns></returns>
        public static string ToUrlQueryString(this Dictionary<string, string> dict, bool urlEncode = true)
        {
            return string.Join("&", dict.Select(p => $"{(urlEncode ? p.Key?.UrlEncode() : "")}={(urlEncode ? p.Value?.UrlEncode() : "")}"));
        }

        /// <summary>
        /// 移除空值项
        /// </summary>
        /// <param name="dict"></param>
        public static void RemoveEmptyValueItems(this Dictionary<string, string> dict)
        {
            dict.Where(item => string.IsNullOrEmpty(item.Value)).Select(item => item.Key).ToList().ForEach(key =>
            {
                dict.Remove(key);
            });
        }

    }
}
