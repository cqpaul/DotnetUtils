using Cqpaul.Dotnet.Util.Extensions;
using Cqpaul.Dotnet.Util.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Cqpaul.Dotnet.Util.Test.HelpersTest
{
    public class PhotoHelperTest
    {
        [Fact]
        public void GetPhotoMetaInfosTest()
        {
            var filePath = @"C:\Users\pauzhang\Desktop\MPaul\P70513-181220.jpg";
            var result = PhotoHelper.GetPhotoMetaInfos(filePath);
            Assert.NotNull(result);
        }

        [Fact]
        public void GetPhotoXYInfoTest()
        {
            var filePath = @"C:\Users\pauzhang\Desktop\MPaul\P70513-181220.jpg";
            var result = PhotoHelper.GetPhotoXYInfo(filePath);
            Assert.NotNull(result);
        }

    }
}
