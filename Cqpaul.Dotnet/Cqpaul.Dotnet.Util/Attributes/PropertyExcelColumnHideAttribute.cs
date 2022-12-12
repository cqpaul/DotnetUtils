using System;

namespace Cqpaul.Dotnet.Util.Attributes
{
    /// <summary>
    /// 类型字段与Excel Column名的映射关系
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class PropertyExcelColumnHideAttribute : Attribute
    {
        public bool IsHiding;

        public PropertyExcelColumnHideAttribute(bool isHiding)
        {
            this.IsHiding = isHiding;
        }

    }
}
