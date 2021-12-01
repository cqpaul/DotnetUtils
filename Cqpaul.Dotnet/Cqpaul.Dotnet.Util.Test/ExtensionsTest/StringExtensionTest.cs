using Cqpaul.Dotnet.Util.Extensions;
using System;
using Xunit;

namespace Cqpaul.Dotnet.Util.Test.ExtensionsTest
{
    public class StringExtensionTest
    {
        [Fact]
        public void AppendTimestampForFileNameTest()
        {
            string fileName = "StringExtensionTest.cs";
            string updateFileName = fileName.AppendStrToFileName(Guid.NewGuid().ToString("N"), "_");
            Assert.NotEmpty(updateFileName);
        }
    }
}
