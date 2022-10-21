using MetadataExtractor;
using static System.Net.Mime.MediaTypeNames;

namespace Cqpaul.Dotnet.Util.Helpers
{
    public static class PhotoHelper
    {
        /// <summary>
        /// 读取文件的元数据信息
        /// </summary>
        /// <param name="filePath"></param>
        public static Dictionary<string, List<string?>> GetPhotoMetaInfos(string filePath)
        {
            Dictionary<string, List<string?>> photoMetaInfos = new Dictionary<string, List<string?>>();
            IReadOnlyList<MetadataExtractor.Directory> metaDrectories = ImageMetadataReader.ReadMetadata(filePath);
            foreach (var directory in metaDrectories)
            {
                foreach (Tag tag in directory.Tags)
                {
                    if (!photoMetaInfos.ContainsKey(tag.Name))
                    {
                        photoMetaInfos.Add(tag.Name, new List<string?>());
                    }
                    photoMetaInfos[tag.Name].Add(tag.Description);
                }
            }
            return photoMetaInfos;
        }

        /// <summary>
        /// 读取图片的经纬度信息
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static (string? Longitude, string? Latitude)? GetPhotoXYInfo(string filePath)
        {
            var allMetaInfo = PhotoHelper.GetPhotoMetaInfos(filePath);
            if (allMetaInfo.Any(a => a.Key == "GPS Longitude"))
            {
                return (allMetaInfo["GPS Longitude"].First(), allMetaInfo["GPS Latitude"].First());
            }
            return null;
        }

    }
}
