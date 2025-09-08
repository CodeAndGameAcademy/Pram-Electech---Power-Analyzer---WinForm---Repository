using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerAnalyzer.Util
{
    public class Helper
    {
        public static byte[] HexStringToByteArray(string hex)
        {
            // Remove spaces and dashes
            hex = hex.Replace(" ", "").Replace("-", "");

            // Must be even length
            if (hex.Length % 2 != 0)
                throw new ArgumentException("Invalid hex string length.");

            byte[] bytes = new byte[hex.Length / 2];

            for (int i = 0; i < bytes.Length; i++)
            {
                bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
            }

            return bytes;
        }

        public static string HexStringToDecimal(string hexString, string hexDp)
        {
            int value = Convert.ToInt32(hexString, 16);

            // Convert hex string to decimal places (e.g., "06" -> 6)
            int decimalPlaces = int.Parse(hexDp, System.Globalization.NumberStyles.HexNumber);

            // Scale the value
            double result = value / Math.Pow(10, decimalPlaces);

            // Format to fixed decimal places (no E notation)
            return result.ToString("F" + decimalPlaces);
        }

        // Helper method: Find byte sequence in list
        public static int FindSequence(List<byte> buffer, byte[] sequence, int startIndex = 0)
        {
            for (int i = startIndex; i <= buffer.Count - sequence.Length; i++)
            {
                bool match = true;
                for (int j = 0; j < sequence.Length; j++)
                {
                    if (buffer[i + j] != sequence[j])
                    {
                        match = false;
                        break;
                    }
                }
                if (match) return i;
            }
            return -1;
        }
    }
}
