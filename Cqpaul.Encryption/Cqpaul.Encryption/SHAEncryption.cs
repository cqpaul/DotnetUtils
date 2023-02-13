namespace Cqpaul.Encryption
{
    /// <summary>
    /// SHA，全称SecureHashAlgorithm，是一种数据加密算法，该算法的思想是接收一段明文，
    /// 然后以一种不可逆的方式将它转换成一段（通常更小）密文，也可以简单的理解为取一串输入码（称为预映射或信息），
    /// 并把它们转化为长度较短、位数固定的输出序列即散列值（也称为信息摘要或信息认证代码）的过程。
    /// SHA为不可逆加密方式。
    /// </summary>
    public static class SHAEncryption
    {
        public static string SHA1Encrypt(string Txt)
        {
            var bytes = System.Text.Encoding.Default.GetBytes(Txt);
            var SHA = System.Security.Cryptography.SHA1.Create();
            var encryptbytes = SHA.ComputeHash(bytes);
            return Convert.ToBase64String(encryptbytes);
        }
        
        public static string SHA256Encrypt(string Txt)
        {
            var bytes = System.Text.Encoding.Default.GetBytes(Txt);
            var SHA256 = System.Security.Cryptography.SHA256.Create();
            var encryptbytes = SHA256.ComputeHash(bytes);
            return Convert.ToBase64String(encryptbytes);
        }

        public static string SHA384Encrypt(string Txt)
        {
            var bytes = System.Text.Encoding.Default.GetBytes(Txt);
            var SHA384 = System.Security.Cryptography.SHA384.Create();
            var encryptbytes = SHA384.ComputeHash(bytes);
            return Convert.ToBase64String(encryptbytes);
        }

        public static string SHA512Encrypt(string Txt)
        {
            var bytes = System.Text.Encoding.Default.GetBytes(Txt);
            var SHA512 = System.Security.Cryptography.SHA512.Create();
            var encryptbytes = SHA512.ComputeHash(bytes);
            return Convert.ToBase64String(encryptbytes);
        }
    }
}
