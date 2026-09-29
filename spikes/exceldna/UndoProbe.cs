using System;
using System.Collections.Generic;
using System.Reflection;

namespace EmtSpike
{
    /// <summary>K2: reads Excel's native undo list without changing anything.</summary>
    internal static class UndoProbe
    {
        public static object Snapshot(string phase, string variant)
        {
            dynamic app = Xl.App;
            object commandBars = app.CommandBars;
            object enabled = null;
            int count = -1;
            string source = null;
            var items = new List<string>();
            var errors = new List<string>();

            try { enabled = app.CommandBars.GetEnabledMso("Undo"); }
            catch (Exception ex) { errors.Add("GetEnabledMso(Undo): " + ex.Message); }

            // Source 1: CommandBars("Standard").Controls("&Undo"); source 2: FindControl(Id:=128).
            var sources = new List<(string Name, Func<object> Get)>
            {
                ("Standard/&Undo", () => Xl.Get(Xl.Get(Xl.Get(commandBars, "Item", "Standard"), "Controls"), "Item", "&Undo")),
                ("FindControl(Id:=128)", () => commandBars.GetType().InvokeMember("FindControl", BindingFlags.InvokeMethod,
                    null, commandBars, new object[] { 128 }, null, null, new[] { "Id" })),
            };
            foreach (var (name, get) in sources)
            {
                try
                {
                    object ctrl = get();
                    if (ctrl == null) { errors.Add(name + ": not found"); continue; }
                    count = Convert.ToInt32(Xl.Get(ctrl, "ListCount"));
                    for (int i = 1; i <= count; i++)
                        items.Add(Convert.ToString(Xl.Get(ctrl, "List", i)));
                    source = name;
                    break;
                }
                catch (Exception ex)
                {
                    errors.Add(name + ": " + (ex.InnerException ?? ex).Message);
                }
            }

            Log.Write("undo", new { phase, variant, enabled, count, items, source, errors });
            return new { phase, enabled, count, top = items.Count > 0 ? items[0] : null };
        }
    }
}
