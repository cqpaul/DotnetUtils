namespace Cqpaul.Dotnet.Util.Enums
{
    /// <summary>
    /// Excel 数据 需要校验的类型
    /// </summary>
    public enum ExcelDataValidateType
    {
        /// <summary>
        /// 不需要校验
        /// </summary>
        DoNotNeedValidate,
        /// <summary>
        /// 必填
        /// </summary>
        MustHave,
        /// <summary>
        /// 必须数字，含百分数
        /// </summary>
        MustNumber,
        /// <summary>
        /// 必须日期
        /// </summary>
        MustDate,
        /// <summary>
        /// 必须在选项之中
        /// </summary>
        MustInOptions,
        /// <summary>
        /// 如有值，则必须数字，含百分数
        /// </summary>
        IfNotNullMustNumber,
        /// <summary>
        /// 如有值，则必须日期
        /// </summary>
        IfNotNullMustDate,
        /// <summary>
        /// 如有值，则必须在选项之中
        /// </summary>
        IfNotNullMustInOptions,
        /// <summary>
        /// 值为‘未支付’或日期格式
        /// </summary>
        MustDateOrString,
        /// <summary>
        /// 值为‘未收到’或日期格式
        /// </summary>
        MustDateOrWString,

    }
}
