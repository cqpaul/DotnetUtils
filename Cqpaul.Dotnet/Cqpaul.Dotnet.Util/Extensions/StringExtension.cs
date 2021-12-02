using System.Text;

namespace Cqpaul.Dotnet.Util.Extensions
{
    public static class StringExtension
    {
        /// <summary>
        /// 在文件名中，Append相关字符串
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="appendStr"></param>
        /// <param name="Separator"></param>
        /// <returns></returns>
        public static string AppendStrToFileName(this string fileName, string appendStr, string Separator)
        {
            List<string> nameParts = fileName.Split('.').ToList();
            string fileSuffix = nameParts.Last();
            nameParts.RemoveAt(nameParts.Count() - 1);
            string prefixFileName = string.Join(".", nameParts);
            return $"{prefixFileName}{Separator}{appendStr}.{fileSuffix}";
        }

        /// <summary>
        /// 将一个字符串 URL 编码
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string UrlEncode(this string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return "";
            }
            return System.Web.HttpUtility.UrlEncode(str, Encoding.UTF8);
        }

    }
}
