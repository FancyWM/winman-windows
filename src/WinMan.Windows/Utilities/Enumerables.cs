using System;
using System.Collections.Generic;
using System.Linq;

namespace WinMan.Windows.Utilities
{
    internal static class Enumerables
    {
        public static (List<TDerived> added, List<TDerived> removed) UpdateDerivedObjectListUnderLock<TDerived, TBase, TKey>(object syncRoot, List<TDerived> derivedList, Func<TDerived, TKey> derivedKeySelector, List<TBase> baseList, Func<TBase, TKey> baseKeySelector, Func<TBase, TDerived> produce)
                where TKey : IEquatable<TKey>
        {
            List<TKey> baseKeys = [.. baseList.Select(baseKeySelector)];
            List<TDerived> added = [];
            List<TDerived> removed = [];

            lock (syncRoot)
            {
                List<TDerived> oldDerivedList = [.. derivedList];
                if (oldDerivedList.Select(derivedKeySelector).SequenceEqual(baseKeys))
                {
                    return (added, removed);
                }

                derivedList.Clear();
                for (int i = 0; i < baseKeys.Count; i++)
                {
                    var baseKey = baseKeys[i];
                    var preservedDerived = oldDerivedList.FirstOrDefault(x => derivedKeySelector(x).Equals(baseKey));
                    if (preservedDerived != null)
                    {
                        derivedList.Add(preservedDerived);
                    }
                    else
                    {
                        var newBase = baseList.First(x => baseKeySelector(x).Equals(baseKey));
                        try
                        {
                            var newDerived = produce(newBase);
                            derivedList.Add(newDerived);
                            added.Add(newDerived);
                        }
                        catch
                        {
                            // Ignore...
                        }
                    }
                }

                removed = [.. oldDerivedList.Where(x => !derivedList.Contains(x))];
            }

            return (added, removed);
        }
    }
}
