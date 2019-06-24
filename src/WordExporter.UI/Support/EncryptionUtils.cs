using System;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Cryptography;
using System.Text;

namespace WordExporter.UI.Support
{
    public static class EncryptionUtils
    {
        public static String Encrypt(String data)
        {
            var encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(data), null, DataProtectionScope.CurrentUser);
            return System.Convert.ToBase64String(encrypted);
        }

        public static String Encrypt(SecureString secureString)
        {
            var stringToEncrypt = SecureStringToString(secureString);
            var encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(stringToEncrypt), null, DataProtectionScope.CurrentUser);
            return System.Convert.ToBase64String(encrypted);
        }

        public static String Decrypt(String encryptedData)
        {
            if (String.IsNullOrEmpty(encryptedData))
                return "";

            try
            {
                var decrypted = ProtectedData.Unprotect(
                    Convert.FromBase64String(encryptedData),
                    null,
                    DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(decrypted);
            }
            catch (Exception e)
            {
                //TODO logging decrypting data.
                return "";
            }
        }

        public static String SecureStringToString(SecureString value)
        {
            IntPtr valuePtr = IntPtr.Zero;
            try
            {
                valuePtr = Marshal.SecureStringToGlobalAllocUnicode(value);
                return Marshal.PtrToStringUni(valuePtr);
            }
            finally
            {
                Marshal.ZeroFreeGlobalAllocUnicode(valuePtr);
            }
        }
    }
}
