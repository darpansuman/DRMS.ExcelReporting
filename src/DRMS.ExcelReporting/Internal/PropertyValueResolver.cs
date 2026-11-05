using System.Collections;
using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;

namespace DRMS.ExcelReporting.Internal;

/// <summary>
/// Reads a value from an object using a property path such as "Address.City".
/// </summary>
/// <remarks>
/// - Matching ignores upper or lower case.
/// - Public properties and public fields are supported.
/// - Dictionaries (for example ExpandoObject or Dictionary&lt;string, object&gt;) are supported, so the
///   generator also works with dynamic data.
/// - If any step in the path is null, the result is null (no exception).
/// - Getters are compiled once per type and member, then cached, so large lists stay fast.
/// </remarks>
internal static class PropertyValueResolver
{
    private static readonly ConcurrentDictionary<(Type Type, string Member), Func<object, object?>?> GetterCache = new();

    /// <summary>
    /// Returns the value at <paramref name="propertyPath"/>, or null when a part of the path is null.
    /// Throws <see cref="InvalidOperationException"/> when a property name does not exist.
    /// </summary>
    public static object? GetValue(object? source, string propertyPath)
    {
        var current = source;

        foreach (var member in propertyPath.Split('.', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
        {
            if (current is null)
            {
                return null;
            }

            current = GetMemberValue(current, member);
        }

        return current;
    }

    private static object? GetMemberValue(object source, string member)
    {
        // Dynamic data: read the key from the dictionary.
        if (source is IDictionary<string, object?> dictionary)
        {
            return TryGetDictionaryValue(dictionary, member);
        }

        if (source is IDictionary legacyDictionary)
        {
            return legacyDictionary.Contains(member) ? legacyDictionary[member] : null;
        }

        var getter = GetterCache.GetOrAdd((source.GetType(), member), static key => BuildGetter(key.Type, key.Member));

        if (getter is null)
        {
            throw new InvalidOperationException(
                $"Type '{source.GetType().Name}' has no public property or field named '{member}'. " +
                "Check the \"Property\" value of the column in appsettings.json.");
        }

        return getter(source);
    }

    private static object? TryGetDictionaryValue(IDictionary<string, object?> dictionary, string key)
    {
        if (dictionary.TryGetValue(key, out var value))
        {
            return value;
        }

        // Fall back to a case-insensitive search.
        foreach (var pair in dictionary)
        {
            if (string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase))
            {
                return pair.Value;
            }
        }

        return null;
    }

    /// <summary>
    /// Builds a fast delegate: (object source) => (object)((TSource)source).Member
    /// Returns null when the member does not exist.
    /// </summary>
    private static Func<object, object?>? BuildGetter(Type type, string memberName)
    {
        const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase;

        // Ignore indexers such as this[int]; they need arguments.
        MemberInfo? member =
            type.GetProperties(flags).FirstOrDefault(p =>
                p.GetIndexParameters().Length == 0 &&
                p.CanRead &&
                string.Equals(p.Name, memberName, StringComparison.OrdinalIgnoreCase))
            ?? (MemberInfo?)type.GetField(memberName, flags);

        if (member is null)
        {
            return null;
        }

        var parameter = Expression.Parameter(typeof(object), "source");
        var typedSource = Expression.Convert(parameter, type);
        var access = Expression.MakeMemberAccess(typedSource, member);
        var boxed = Expression.Convert(access, typeof(object));

        return Expression.Lambda<Func<object, object?>>(boxed, parameter).Compile();
    }
}
