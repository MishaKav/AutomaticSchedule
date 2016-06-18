using System.Linq;
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Text;

namespace AutomaticSchedule
{
    public static class ExtensionMethods
    {
        #region Default Extentions

        public static bool IsNullOrEmpty(this string source)
        {
            return string.IsNullOrEmpty(source);
        }

        public static bool IsNotNullOrEmpty(this string source)
        {
            return !string.IsNullOrEmpty(source);
        }

        public static bool IsNotEmptyObject(this object prop)
        {
            return !prop.IsEmptyObject();
        }

        public static bool IsEmptyObject(this object prop)
        {
            if (prop == null)
            {
                return true;
            }
            else if (prop is string)
            {
                return prop.ToString().Length == 0;
            }
            else
            {
                var ps = prop.GetType().GetProperties();

                foreach (var pi in ps)
                {
                    var value = pi.GetValue(prop, null);
                    var valueStr = value?.ToString() ?? string.Empty;    // if value not a class inside class

                    if (valueStr.IsNotNullOrEmpty())
                        return false;
                }

            }
            return true;
        }

        public static bool IsAny<T>(this IEnumerable<T> source)
        {
            return source != null && source.Any();
        }

        public static string ToFullFormat(this DateTime source)
        {
            return source.ToString("dd/MM/yyyy hh:MM");
        }

        public static string ToReadableFileSize(this long source)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            var order = 0;
            double result = source;
            while (result >= 1024 && order + 1 < sizes.Length)
            {
                order++;
                result = result / 1024;
            }

            if (order == 0)
            {
                result = 1;
                order++;
            }
            // Adjust the format string to your preferences. For example "{0:0.#}{1}" would
            // show a single decimal place, and no space.
            return $"{result:0.##} {sizes[order]}";
        }

        public static T ToEnum<T>(this string value)
        {
            if (value.IsNullOrEmpty())
            {
                return default(T);
            }

            try
            {
                var result = (T)Enum.Parse(typeof(T), value, true);
                return result;
            }
            catch
            {
                return default(T);
            }
        }

        public static bool EqualsIgnoreCase(this string source, string value)
        {
            if (source.IsNullOrEmpty())
            {
                return false;
            }

            return source.Equals(value, StringComparison.CurrentCultureIgnoreCase);
        }

        public static bool ContainsIgnoreCase(this string source, string value)
        {
            if (source.IsNullOrEmpty())
            {
                return false;
            }

            return source.IndexOf(value, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        public static string EncodingUtf8(this string source)
        {
            if (source.IsNullOrEmpty())
            {
                return string.Empty;
            }

            var bytes = Encoding.Default.GetBytes(source);
            return Encoding.UTF8.GetString(bytes);
        }

        public static string ToDefaultFormat(this Stopwatch source)
        {
            if (source != null)
            {
                return source.Elapsed.TotalSeconds > 60 ? 
                    $"{source.Elapsed.TotalMinutes.ToString("N0")}m {source.Elapsed.Seconds.ToString("N0")}s" : 
                    $"{source.Elapsed.TotalSeconds.ToString("N2")}s";
            }

            return string.Empty;
        }


        #endregion Default Extentions

        #region Global Utils for DateTime

        public static DateTime ChangeTime(this DateTime dateTime, int hours, int minutes, int seconds = 0, int milliseconds = 0)
        {
            return new DateTime(
                dateTime.Year,
                dateTime.Month,
                dateTime.Day,
                hours,
                minutes,
                seconds,
                milliseconds,
                dateTime.Kind);
        }

        public static string ToDefaultDateFormat(this DateTime source)
        {
            return source.ToString("dd/MM/yyyy");
        }

        public static string ToDefaultTimeFormat(this DateTime source)
        {
            return source.ToString("HH:mm");
        }

        public static string ToDefaultDateTimeFormat(this DateTime source)
        {
            return $"{source.ToDefaultDateFormat()} {source.ToDefaultTimeFormat()}";
        }

        public static string ToDefaultDurationFormat(this TimeSpan source)
        {
            return source.Days > 0 ?
                $"{source.Days}D {source.Hours}H {source.Minutes}M" :
                $"{source.Hours}H {source.Minutes}M";
        }

        public static string ToDefaultDurationFormat(this double source)
        {
            var ts = TimeSpan.FromMinutes(source);
            return ts.Days > 0 ?
                $"{ts.Days}D {ts.Hours}H {ts.Minutes}M" :
                $"{ts.Hours}H {ts.Minutes}M";
        }

        public static string ToMoneyFormat(this double source)
        {
            return Math.Ceiling(source).ToString("#,#");
        }

        public static string ToEstimateDuration(this string minutes)
        {
            if (minutes.IsNullOrEmpty()) return "0 Min";

            var minutesDouble = double.Parse(minutes);
            return minutesDouble.ToEstimateDuration();
        }

        public static string ToEstimateDuration(this double minutes)
        {
            var ts = TimeSpan.FromMinutes(minutes);
            return ts.Hours > 0 ? $"{ts.Hours} Hrs {ts.Minutes} Min" : $"{ts.Minutes} Min";
        }

        #endregion Global Utils

        #region LINQ Extentions

        public static IEnumerable<TSource> DistinctBy<TSource, TKey>(this IEnumerable<TSource> source, Func<TSource, TKey> keySelector)
        {
            var seenKeys = new HashSet<TKey>();
            foreach (var element in source)
            {
                if (seenKeys.Add(keySelector(element)))
                {
                    yield return element;
                }
            }
        }

        #endregion LINQ Extentions
    }
}