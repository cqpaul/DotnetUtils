namespace Cqpaul.Dotnet.Util.Extensions
{
    public static class StringExtension
    {
        public static string AppendStrToFileName(this string fileName, string appendStr, string Separator)
        {
            List<string> nameParts = fileName.Split('.').ToList();
            string fileSuffix = nameParts.Last();
            nameParts.RemoveAt(nameParts.Count() - 1);
            string prefixFileName = string.Join(".", nameParts);
            return $"{prefixFileName}{Separator}{appendStr}.{fileSuffix}";
        }
    }
}
