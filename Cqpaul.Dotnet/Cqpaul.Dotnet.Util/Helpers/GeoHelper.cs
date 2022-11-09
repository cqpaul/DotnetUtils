using MaxRev.Gdal.Core;
using OSGeo.GDAL;
using OSGeo.OGR;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Cqpaul.Dotnet.Util.Helpers
{
    public static class GeoHelper
    {
        /// <summary>
        /// Gdal 之 SHP 转 Geojson
        /// </summary>
        /// <param name="srcPath"></param>
        /// <param name="destPath"></param>
        public static bool Shp2Geojson(string srcPath, string destPath)
        {
            //添加所有配置
            GdalBase.ConfigureAll();
            //配置单个选项
            Gdal.SetConfigOption("GDAL_FILENAME_IS_UTF8", "YES");//支持中文路径
            Gdal.SetConfigOption("DXF_ENCODING", "UTF-8");

            //读取文件，这里自动区别文件类型
            using var ds = Ogr.Open(srcPath, 0);

            //根据文件名创建驱动
            using var dv = Ogr.GetDriverByName("GeoJSON");

            //拷贝数据与转化
            using var ret = dv.CopyDataSource(ds, destPath, null);

            // 手动释放,否则destPath一直被占用
            ret.FlushCache();
            ret.Dispose();
            return true;
        }

    }
}
