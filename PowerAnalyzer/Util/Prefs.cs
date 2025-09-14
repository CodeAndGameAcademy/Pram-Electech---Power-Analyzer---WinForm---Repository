using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerAnalyzer.Util
{
    public static class Prefs
    {
        public static T Get<T>(string key, T defaultValue)
        {
            var value = Properties.Settings.Default[key];
            if (value == null || string.IsNullOrEmpty(value.ToString()))
                return defaultValue;
            return (T)value;
        }

        public static void Set<T>(string key, T value)
        {
            Properties.Settings.Default[key] = value;
            Properties.Settings.Default.Save();
        }
    }
}
