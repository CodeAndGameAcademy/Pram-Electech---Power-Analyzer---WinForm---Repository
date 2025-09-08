using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerAnalyzer.Util
{
    public class HexDecimalConverter
    {
        public static string ToConvert(string hexString, string hexDp)
        {
            int value = Convert.ToInt32(hexString, 16);

            // Convert hex string to decimal places (e.g., "06" -> 6)
            int decimalPlaces = int.Parse(hexDp, System.Globalization.NumberStyles.HexNumber);

            // Scale the value
            double result = value / Math.Pow(10, decimalPlaces);

            // Format to fixed decimal places (no E notation)
            return result.ToString("F" + decimalPlaces);
        }
    }
}
