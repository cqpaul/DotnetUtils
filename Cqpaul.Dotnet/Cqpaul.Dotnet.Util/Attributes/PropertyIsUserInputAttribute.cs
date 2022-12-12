using System;
using System.Collections.Generic;
using System.Text;

namespace Cqpaul.Dotnet.Util.Attributes
{
    /// <summary>
    /// 表达 该属性 是否为用户输入的，还是从之前的模板带出来的
    /// </summary>
    [AttributeUsage(AttributeTargets.Property,AllowMultiple =false, Inherited = true)]
    public class PropertyIsUserInputAttribute : Attribute
    {
        public bool IsUserInput;

        public PropertyIsUserInputAttribute(bool isUserInput)
        {
            IsUserInput = isUserInput;
        }
    }
}
