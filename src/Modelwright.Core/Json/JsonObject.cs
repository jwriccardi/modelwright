using System;
using System.Collections.Generic;

namespace Modelwright.Core.Json;

/// <summary>
/// A JSON object node: properties in document order, names unique (ordinal). The other nodes are
/// <see cref="List{T}"/> of nodes (arrays), <see cref="string"/>, <see cref="bool"/>, <see cref="JsonNumber"/>
/// and null; <see cref="JsonWriter"/> also accepts <see cref="int"/>.
/// </summary>
internal sealed class JsonObject
{
    private readonly List<KeyValuePair<string, object?>> _properties = new List<KeyValuePair<string, object?>>();

    // Name -> value, so lookups and duplicate checks are O(1); _properties keeps the document order.
    private readonly Dictionary<string, object?> _byName = new Dictionary<string, object?>(StringComparer.Ordinal);

    /// <summary>The properties, in document order.</summary>
    public IReadOnlyList<KeyValuePair<string, object?>> Properties => _properties;

    /// <summary>Adds a property.</summary>
    /// <exception cref="ArgumentException">A property with this name exists.</exception>
    public void Add(string name, object? value)
    {
        if (_byName.ContainsKey(name))
        {
            throw new ArgumentException($"Duplicate property \"{name}\".", nameof(name));
        }

        _byName.Add(name, value);
        _properties.Add(new KeyValuePair<string, object?>(name, value));
    }

    /// <summary>True if a property named <paramref name="name"/> exists (ordinal).</summary>
    public bool Contains(string name) => _byName.ContainsKey(name);

    /// <summary>Gets the value of the property named <paramref name="name"/> (ordinal).</summary>
    public bool TryGet(string name, out object? value) => _byName.TryGetValue(name, out value);
}

/// <summary>A JSON number node, kept as its source text so no precision is lost.</summary>
internal sealed class JsonNumber
{
    public JsonNumber(string text) => Text = text;

    /// <summary>The number exactly as written, e.g. <c>10000</c> or <c>1.5e3</c>.</summary>
    public string Text { get; }

    /// <inheritdoc />
    public override string ToString() => Text;
}
