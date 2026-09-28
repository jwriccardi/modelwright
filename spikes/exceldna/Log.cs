using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Web.Script.Serialization;
using ExcelDna.Integration;

namespace EmtSpike
{
    /// <summary>Appends JSON lines to %LOCALAPPDATA%\EmtSpike\log.jsonl. Never throws.</summary>
    internal static class Log
    {
        public static readonly string Dir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmtSpike");
        public static readonly string FilePath = Path.Combine(Dir, "log.jsonl");

        private static readonly object Gate = new object();
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

        public static void Write(string evt, object fields = null)
        {
            try
            {
                var d = new Dictionary<string, object>
                {
                    ["ts"] = DateTime.Now.ToString("o"),
                    ["evt"] = evt,
                };
                if (fields != null)
                    foreach (PropertyInfo p in fields.GetType().GetProperties())
                        d[p.Name] = p.GetValue(fields, null);
                string line = Json.Serialize(d);
                lock (Gate)
                {
                    Directory.CreateDirectory(Dir);
                    File.AppendAllText(FilePath, line + Environment.NewLine);
                }
            }
            catch
            {
                // logging must never break a command
            }
        }

        public static void Error(string where, Exception ex) =>
            Write("error", new { where, type = ex.GetType().Name, message = ex.Message, stack = ex.ToString() });
    }

    /// <summary>Tiny late-bound COM helpers (no interop assembly).</summary>
    internal static class Xl
    {
        public static dynamic App => ExcelDnaUtil.Application;

        public static void Status(string msg)
        {
            try { App.StatusBar = msg; } catch { /* ignore */ }
        }

        /// <summary>Get a (possibly parameterized) COM property via IDispatch.</summary>
        public static object Get(object o, string name, params object[] args) =>
            o.GetType().InvokeMember(name, BindingFlags.GetProperty, null, o, args);

        public static long SelectionCount()
        {
            try { return Convert.ToInt64(App.Selection.CountLarge); } catch { return -1; }
        }

        public static double Ms(System.Diagnostics.Stopwatch sw) => Math.Round(sw.Elapsed.TotalMilliseconds, 3);
    }
}
