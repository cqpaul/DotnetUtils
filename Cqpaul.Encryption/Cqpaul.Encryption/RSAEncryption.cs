﻿namespace Cqpaul.Encryption
{
    /// <summary>
    /// RSA是目前最有影响力的公钥加密算法，多用于数据加密和数字签名。
    /// </summary>
    public static class RSAEncryption
    {
        /// <summary> 
        /// RSA加密数据 
        /// </summary> 
        /// <param name="express">要加密数据</param> 
        /// <param name="KeyContainerName">密匙容器的名称</param> 
        /// <returns></returns> 
        public static string RSAEncrypt(string express, string KeyContainerName = null)
        {

            System.Security.Cryptography.CspParameters param = new System.Security.Cryptography.CspParameters();
            param.KeyContainerName = KeyContainerName ?? "Test"; //密匙容器的名称，保持加密解密一致才能解密成功
            using (System.Security.Cryptography.RSACryptoServiceProvider rsa = new System.Security.Cryptography.RSACryptoServiceProvider(param))
            {
                byte[] plaindata = System.Text.Encoding.Default.GetBytes(express);//将要加密的字符串转换为字节数组
                byte[] encryptdata = rsa.Encrypt(plaindata, false);//将加密后的字节数据转换为新的加密字节数组
                return Convert.ToBase64String(encryptdata);//将加密后的字节数组转换为字符串
            }
        }
        /// <summary> 
        /// RSA解密数据 
        /// </summary> 
        /// <param name="express">要解密数据</param> 
        /// <param name="KeyContainerName">密匙容器的名称</param> 
        /// <returns></returns> 
        public static string RSADecrypt(string ciphertext, string KeyContainerName = null)
        {
            System.Security.Cryptography.CspParameters param = new System.Security.Cryptography.CspParameters();
            param.KeyContainerName = KeyContainerName ?? "Test"; //密匙容器的名称，保持加密解密一致才能解密成功
            using (System.Security.Cryptography.RSACryptoServiceProvider rsa = new System.Security.Cryptography.RSACryptoServiceProvider(param))
            {
                byte[] encryptdata = Convert.FromBase64String(ciphertext);
                byte[] decryptdata = rsa.Decrypt(encryptdata, false);
                return System.Text.Encoding.Default.GetString(decryptdata);
            }
        }
    }
}
