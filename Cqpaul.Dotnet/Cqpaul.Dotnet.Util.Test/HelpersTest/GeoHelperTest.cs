using Cqpaul.Dotnet.Util.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Cqpaul.Dotnet.Util.Test.HelpersTest
{
    public class GeoHelperTest
    {
        [Fact]
        public void Shp2GeojsonTest()
        {
            string sourcePath = @"C:\Users\pauzhang\Desktop\shapetemplate\面图形.shp";
            string destPath = @"C:\Users\pauzhang\Desktop\shapetemplate\out.json";

            var result = GeoHelper.Shp2Geojson(sourcePath, destPath);

            Assert.True(result);
        }
    }
}
