using System;
using System.Security.Cryptography;

namespace DeltaCore.WorkFlow.Utils
{
	public enum RandomPasswordOptions
	{
		Alpha = 1,
		Numeric = 2,
		AlphaNumeric = Alpha + Numeric,
		AlphaNumericSpecial = 4
	}

	public class KeyGenerator
	{
		private const string PASSWORD_CHARS_Alpha = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
		private const string PASSWORD_CHARS_NUMERIC = "0123456789";
		private const string PASSWORD_CHARS_SPECIAL = "*$-+?_!%{}/";
		#region Overloads

		/// <summary>  
		/// Generates a random password with the default length.  
		/// </summary>  
		/// <returns>Randomly generated password.</returns>  
		public string Generate()
		{
			var rdn = new Random(DateTime.Now.Millisecond);
			return Generate(rdn.Next(8, 99), RandomPasswordOptions.AlphaNumericSpecial);
		}

		/// <summary>  
		/// Generates a random password with the default length.  
		/// </summary>  
		/// <returns>Randomly generated password.</returns>  
		public string Generate(RandomPasswordOptions option)
		{
			var rdn = new Random(DateTime.Now.Millisecond);
			return Generate(rdn.Next(8, 99), option);
		}

		/// <summary>  
		/// Generates a random password with the default length.  
		/// </summary>  
		/// <returns>Randomly generated password.</returns>  
		public string Generate(int passwordLength)
		{
			var rdn = new Random(DateTime.Now.Millisecond);
			return Generate(passwordLength, RandomPasswordOptions.AlphaNumericSpecial);
		}

		/// <summary>  
		/// Generates a random password.  
		/// </summary>  
		/// <returns>Randomly generated password.</returns>  
		public string Generate(int passwordLength, RandomPasswordOptions option)
		{
			return GeneratePassword(passwordLength, option);
		}

		#endregion


		/// <summary>  
		/// Generates the password.  
		/// </summary>  
		/// <returns></returns>  
		private string GeneratePassword(int passwordLength, RandomPasswordOptions option)
		{
			if (passwordLength < 0)
				return null;
			var passwordChars = GetCharacters(option);
			if (string.IsNullOrEmpty(passwordChars))
				return null;
			var password = new char[passwordLength];
			var random = GetRandom();
			for (int i = 0; i < passwordLength; i++)
			{
				var index = random.Next(passwordChars.Length);
				var passwordChar = passwordChars[index];
				password[i] = passwordChar;
			}
			return new string(password);
		}

		/// <summary>  
		/// Gets the characters selected by the option  
		/// </summary>  
		/// <returns></returns>  
		private string GetCharacters(RandomPasswordOptions option)
		{
			switch (option)
			{
				case RandomPasswordOptions.Alpha:
					return PASSWORD_CHARS_Alpha;
				case RandomPasswordOptions.Numeric:
					return PASSWORD_CHARS_NUMERIC;
				case RandomPasswordOptions.AlphaNumeric:
					return PASSWORD_CHARS_Alpha + PASSWORD_CHARS_NUMERIC;
				case RandomPasswordOptions.AlphaNumericSpecial:
					return PASSWORD_CHARS_Alpha + PASSWORD_CHARS_NUMERIC +
								 PASSWORD_CHARS_SPECIAL;
			}
			return string.Empty;
		}

		/// <summary>  
		/// Gets a random object with a real random seed  
		/// </summary>  
		/// <returns></returns>  
		private Random GetRandom()
		{
			// Use a 4-byte array to fill it with random bytes and convert it then to an integer value.  
			var randomBytes = new byte[4];
			// Generate 4 random bytes.  
			var rng = new RNGCryptoServiceProvider();
			rng.GetBytes(randomBytes);
			// Convert 4 bytes into a 32-bit integer value.  
			int seed = (randomBytes[0] & 0x7f) << 24 |
								 randomBytes[1] << 16 |
								 randomBytes[2] << 8 |
								 randomBytes[3];
			// Now, this is real randomization.  
			return new Random(seed);
		}
	}
}