using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace DeltaCore.WorkFlow
{
    public class SecurityManager
    {
        #region RSA

        public string RSA_Cifrar(string texto)
        {
            var sec = new RSACryptoServiceProvider();
            byte[] dato = Encoding.UTF8.GetBytes(texto);
            byte[] dato_cifrado = sec.Encrypt(dato, false);
            RSA_Guardarllave(sec.ToXmlString(true));
            return Convert.ToBase64String(dato_cifrado, 0, dato_cifrado.Length);
        }

        public byte[] RSA_Cifrar(byte[] infor, string xmlKeys)
        {
            var rsa = new RSACryptoServiceProvider();
            rsa.FromXmlString(xmlKeys);
            return rsa.Encrypt(infor, false);
        }
        
        public string RSA_DesCifrar(string textocifrado)
        {
            var sec = new RSACryptoServiceProvider();
            byte[] data = Convert.FromBase64String(textocifrado);
            sec.FromXmlString(RSA_Leerllave());
            byte[] datades = sec.Decrypt(data, false);
            return Encoding.UTF8.GetString(datades);
        }

        public byte[] RSA_DesCifrar(byte[] datosEnc, string xmlKeys)
        {
            var rsa = new RSACryptoServiceProvider();
            rsa.FromXmlString(xmlKeys);
            return rsa.Decrypt(datosEnc, false);
        }

        public void RSA_Guardarllave(string llave)
        {
            XmlWriter w = XmlWriter.Create("llave.xml");
            w.WriteComment("Licencia de Activacion");
            w.WriteStartElement("Llave");
            w.WriteElementString("key", llave);
            w.WriteEndElement();
            w.Close();
        }

        public string RSA_Leerllave()
        {
            XmlReader r = XmlReader.Create("llave.xml");
            r.ReadStartElement("Llave");
            string key = r.ReadElementContentAsString();
            r.Close();
            return key;
        }

        public void Generate_KeyRing(string ringPath)
        {
            var rsa = new RSACryptoServiceProvider();
            string xmlKey = rsa.ToXmlString(true);
            // Si no existe el directorio, crearlo
            string dirPruebas = Path.GetDirectoryName(ringPath);
            if (!Directory.Exists(dirPruebas))
            {
                Directory.CreateDirectory(dirPruebas);
            }
            using (var sw = new StreamWriter(ringPath, false, Encoding.UTF8))
            {
                sw.WriteLine(xmlKey);
                sw.Close();
            }
        }

        public string Get_KeyRing(string ringPath)
        {
            string ring;
            using (var sr = new StreamReader(ringPath, Encoding.UTF8))
            {
                ring = sr.ReadToEnd();
                sr.Close();
            }
            return ring;
        }
        
        #endregion

        #region TDES

        public string TDESCifrar(string clave, string cadena)
        {
            //Arreglo donde guardaremos la cadena descifrada.
            byte[] arreglo = Encoding.UTF8.GetBytes(cadena);
            // Ciframos utilizando el Algoritmo MD5.
            var md5 = new MD5CryptoServiceProvider();
            byte[] llave = md5.ComputeHash(Encoding.UTF8.GetBytes(clave));
            md5.Clear();
            //Ciframos utilizando el Algoritmo 3DES.
            var tripledes = new TripleDESCryptoServiceProvider
                                {Key = llave, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7};
            // Iniciamos la conversión de la cadena
            ICryptoTransform convertir = tripledes.CreateEncryptor();
            //Arreglo de bytes donde guardaremos la cadena cifrada.  
            byte[] resultado = convertir.TransformFinalBlock(arreglo, 0, arreglo.Length);
            tripledes.Clear();
            // Convertimos la cadena y la regresamos.
            return Convert.ToBase64String(resultado, 0, resultado.Length);
        }

        public string TDESDescifrar(string clave, string cadena)
        {
            byte[] arreglo = Convert.FromBase64String(cadena); // Arreglo donde guardaremos la cadena descovertida.
            // Ciframos utilizando el Algoritmo MD5.
            var md5 = new MD5CryptoServiceProvider();
            byte[] llave = md5.ComputeHash(Encoding.UTF8.GetBytes(clave));
            md5.Clear();
            //Ciframos utilizando el Algoritmo 3DES.
            var tripledes = new TripleDESCryptoServiceProvider
                                {Key = llave, Mode = CipherMode.ECB, Padding = PaddingMode.PKCS7};
            ICryptoTransform convertir = tripledes.CreateDecryptor();
            byte[] resultado = convertir.TransformFinalBlock(arreglo, 0, arreglo.Length);
            tripledes.Clear();
            // Obtenemos la cadena y se devuelve
            return Encoding.UTF8.GetString(resultado);
        }

        #endregion

        #region AES

        private void ValidarArgumentos(byte[] Key, byte[] IV)
        {
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("Key");
        }

        public byte[] AESencrypt(string plainText, byte[] Key, byte[] IV)
        {
            if (string.IsNullOrEmpty(plainText))
                throw new ArgumentNullException("plainText");
            ValidarArgumentos(Key, IV);
            MemoryStream msEncrypt = null;
            CryptoStream csEncrypt = null;
            StreamWriter swEncrypt = null;
            Aes aesAlg = null;
            try
            {
                aesAlg = Aes.Create();
                aesAlg.Key = Key;
                aesAlg.IV = IV;
                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                msEncrypt = new MemoryStream();
                csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write);
                swEncrypt = new StreamWriter(csEncrypt);
                swEncrypt.Write(plainText);
            }
            finally
            {
                if (swEncrypt != null) swEncrypt.Close();
                if (csEncrypt != null) csEncrypt.Close();
                if (msEncrypt != null) msEncrypt.Close();
                if (aesAlg != null) aesAlg.Clear();
            }
            return msEncrypt.ToArray();
        }

        public string AESdecrypt(byte[] cipherText, byte[] Key, byte[] IV)
        {
            if (cipherText == null || cipherText.Length <= 0)
                throw new ArgumentNullException("cipherText");
            ValidarArgumentos(Key, IV);
            MemoryStream msDecrypt = null;
            CryptoStream csDecrypt = null;
            StreamReader srDecrypt = null;
            Aes aesAlg = null;
            string plaintext = null;
            try
            {
                aesAlg = Aes.Create();
                aesAlg.Key = Key;
                aesAlg.IV = IV;
                ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                msDecrypt = new MemoryStream(cipherText);
                csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read);
                srDecrypt = new StreamReader(csDecrypt);
                plaintext = srDecrypt.ReadToEnd();
            }
            finally
            {
                if (srDecrypt != null) srDecrypt.Close();
                if (csDecrypt != null) csDecrypt.Close();
                if (msDecrypt != null) msDecrypt.Close();
                if (aesAlg != null) aesAlg.Clear();
            }
            return plaintext;
        }

        #endregion        
    }
}