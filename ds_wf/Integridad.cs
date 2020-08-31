using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace DeltaCore.WorkFlow
{
    public class Integridad : IDisposable
    {
        #region References

        private Boolean disposed;

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            //if (!this.disposed)
            //if (disposing)
            this.disposed = true;
        }

        ~Integridad()
        {
            this.Dispose(false);
        }

        #endregion

        public string DigitoVerificador(string r)
        {
            int suma = 0;
            for (int x = r.Length - 1; x >= 0; x--)
                suma += Int32.Parse(Char.IsDigit(r[x]) ? r[x].ToString() : "0")*(((r.Length - (x + 1))%6) + 2);
            int numericDigito = (11 - suma%11);
            string digito = numericDigito == 11 ? "0" : numericDigito == 10 ? "K" : numericDigito.ToString();
            return digito;
        }

        public bool FilesAreEqual_OneByte83322375(FileInfo first, FileInfo second)
        {
            if (first.Length != second.Length)
                return false;

            using (FileStream fs1 = first.OpenRead())
            using (FileStream fs2 = second.OpenRead())
            {
                for (int i = 0; i < first.Length; i++)
                {
                    if (fs1.ReadByte() != fs2.ReadByte())
                        return false;
                }
            }

            return true;
        }

        public bool FilesAreEqual_Hash1045463924(FileInfo first, FileInfo second)
        {
            byte[] firstHash = MD5.Create().ComputeHash(first.OpenRead());
            byte[] secondHash = MD5.Create().ComputeHash(second.OpenRead());
            return !firstHash.Where((t, i) => t != secondHash[i]).Any();
        }

        public bool FilesAreEqual_OneByte(FileInfo first, FileInfo second)
        {
            if (first.Length != second.Length)
                return false;

            using (FileStream fs1 = first.OpenRead())
            using (FileStream fs2 = second.OpenRead())
            {
                for (int i = 0; i < first.Length; i++)
                {
                    if (fs1.ReadByte() != fs2.ReadByte())
                        return false;
                }
            }

            return true;
        }

        public bool FilesAreEqual_Hash(FileInfo first, FileInfo second)
        {
            byte[] firstHash = MD5.Create().ComputeHash(first.OpenRead());
            byte[] secondHash = MD5.Create().ComputeHash(second.OpenRead());
            return !firstHash.Where((t, i) => t != secondHash[i]).Any();
        }

        public string GetHash(string texto)
        {
            var enc = new ASCIIEncoding();
            var hash = SHA256.Create();
            byte[] checksum = hash.ComputeHash(enc.GetBytes(texto));
            return BitConverter.ToString(checksum).Replace("-", String.Empty);
        }

        public string GetSHA1(string str)
        {
            SHA1 sha1 = SHA1.Create();
            var encoding = new ASCIIEncoding();
            byte[] stream = null;
            var sb = new StringBuilder();
            stream = sha1.ComputeHash(encoding.GetBytes(str));
            foreach (byte t in stream)
                sb.AppendFormat("{0:x2}", t);
            return sb.ToString();
        }

        public string GetCheckSum_MD5(string filename)
        {
            if (!File.Exists(filename))
                return string.Empty;
            using (FileStream stream = File.OpenRead(filename))
            {
                byte[] checksum;
                var md5 = MD5.Create();
                checksum = md5.ComputeHash(stream);
                return BitConverter.ToString(checksum).Replace("-", String.Empty);
            }
        }

        public string GetCheckSum_SHA256(string filename)
        {
            if (!File.Exists(filename))
                throw new ArgumentException("No hay un en la ruta");
            StringBuilder formatted;
            using (var fs = new FileStream(filename, FileMode.Open))
            {
                using (var bs = new BufferedStream(fs))
                {
                    using (var sha = new SHA256Managed())
                    {
                        byte[] hash = sha.ComputeHash(bs);
                        formatted = new StringBuilder(2*hash.Length);
                        foreach (byte b in hash)
                        {
                            formatted.AppendFormat("{0:X2}", b);
                        }
                    }
                }
            }
            return formatted.ToString();
        }
    }
}