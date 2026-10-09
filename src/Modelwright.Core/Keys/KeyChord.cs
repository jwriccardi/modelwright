using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Modelwright.Core.Keys;

/// <summary>
/// A keyboard shortcut: one or more modifiers plus one key, e.g. <c>Ctrl+Shift+[</c>.
/// Parsed from a human key string and rendered to the Excel <c>OnKey</c> / <c>xlcOnKey</c> syntax.
/// </summary>
/// <remarks>
/// <para>
/// Keys are letters <c>A</c>-<c>Z</c>, digits <c>0</c>-<c>9</c>, any printable ASCII punctuation character,
/// or a named key: <c>Up Down Left Right PgUp PgDn Home End Ins Del F1</c>-<c>F12</c>
/// (aliases <c>PageUp PageDown PgDown Insert Delete</c> are accepted).
/// Modifiers are <c>Ctrl</c> (alias <c>Control</c>), <c>Alt</c> and <c>Shift</c>, case-insensitive, in any order.
/// </para>
/// <para>
/// <c>Ctrl+Shift+\</c> and <c>Ctrl+|</c> are different chords, as they are to Excel's <c>OnKey</c>
/// (<c>^+\</c> and <c>^|</c>), but on a US keyboard they are pressed with the same keys;
/// <see cref="ToUsKeys"/> finds such pairs. The <c>+</c> key itself is written <c>Ctrl++</c>.
/// </para>
/// </remarks>
public sealed class KeyChord : IEquatable<KeyChord>
{
    /// <summary>Printable ASCII punctuation accepted as a key.</summary>
    private const string Punctuation = "`-=[]\\;',./~!@#$%^&*()_+{}|:\"<>?";

    /// <summary>Characters that <c>OnKey</c> requires inside braces, because they are otherwise syntax.</summary>
    private const string BracedInOnKey = "[]{}()+^%~";

    /// <summary>US keyboard shifted punctuation, and (same position in <see cref="UsUnshifted"/>) its key.</summary>
    private const string UsShifted = "~!@#$%^&*()_+{}|:\"<>?";

    /// <summary>The unshifted character of the key that types each <see cref="UsShifted"/> character.</summary>
    private const string UsUnshifted = "`1234567890-=[]\\;',./";

    /// <summary>Canonical named key (display form) -> OnKey code.</summary>
    private static readonly Dictionary<string, string> NamedKeys = CreateNamedKeys();

    /// <summary>Accepted spelling (case-insensitive) -> canonical named key.</summary>
    private static readonly Dictionary<string, string> NamedKeyAliases = CreateNamedKeyAliases();

    private KeyChord(KeyModifiers modifiers, string key)
    {
        Modifiers = modifiers;
        Key = key;
    }

    /// <summary>The modifiers (never <see cref="KeyModifiers.None"/>).</summary>
    public KeyModifiers Modifiers { get; }

    /// <summary>
    /// The key in canonical display form: an upper-case letter, a digit, a punctuation character,
    /// or a named key such as <c>Up</c>, <c>PgUp</c>, <c>Ins</c> or <c>F2</c>.
    /// </summary>
    public string Key { get; }

    /// <summary>True when <see cref="Key"/> is a named key (arrow, paging, Ins/Del or function key).</summary>
    public bool IsNamedKey => NamedKeys.ContainsKey(Key);

    /// <summary>
    /// Parses a human key string such as <c>Ctrl+Shift+[</c>, <c>Ctrl+'</c>, <c>Alt+Shift+PgUp</c> or <c>Ctrl+F2</c>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is null.</exception>
    /// <exception cref="FormatException">
    /// The text is empty, has no modifier, an unknown or duplicate modifier, an unknown key,
    /// or is Shift alone with a character key (which would block typing that character).
    /// </exception>
    public static KeyChord Parse(string text)
    {
        if (text is null)
        {
            throw new ArgumentNullException(nameof(text));
        }

        var s = text.Trim();
        if (s.Length == 0)
        {
            throw Invalid(text, "it is empty.");
        }

        string keyToken;
        string? modifierPart;
        if (s.EndsWith("+", StringComparison.Ordinal))
        {
            // The key is '+' itself ("Ctrl++"); a separator '+' must precede it.
            var rest = s.Substring(0, s.Length - 1).TrimEnd();
            if (rest.Length == 0)
            {
                throw Invalid(text, "it has no modifier. Use Ctrl, Alt and/or Shift, e.g. \"Ctrl++\".");
            }

            if (!rest.EndsWith("+", StringComparison.Ordinal))
            {
                throw Invalid(text, "it has no key after the last '+'.");
            }

            keyToken = "+";
            modifierPart = rest.Substring(0, rest.Length - 1);
        }
        else
        {
            var separator = s.LastIndexOf('+');
            keyToken = s.Substring(separator + 1).Trim();
            modifierPart = separator < 0 ? null : s.Substring(0, separator);
        }

        var key = ParseKey(text, keyToken);
        var modifiers = ParseModifiers(text, modifierPart);

        if (modifiers == KeyModifiers.Shift && !NamedKeys.ContainsKey(key))
        {
            throw Invalid(text, "Shift alone with a character key would block typing that character. Add Ctrl or Alt.");
        }

        return new KeyChord(modifiers, key);
    }

    /// <summary>
    /// Renders the chord in Excel <c>OnKey</c> / <c>xlcOnKey</c> syntax: <c>^</c> Ctrl, <c>%</c> Alt, <c>+</c> Shift
    /// (always in that order), letters in lower case, <c>[ ] { } ( ) + ^ % ~</c> and named keys in braces.
    /// Examples: <c>^+{[}</c>, <c>^'</c>, <c>^%\</c>, <c>%+;</c>, <c>^%+{UP}</c>, <c>^{F2}</c>, <c>^+k</c>.
    /// </summary>
    public string ToOnKeyString()
    {
        var sb = new StringBuilder(12);
        if ((Modifiers & KeyModifiers.Ctrl) != 0)
        {
            sb.Append('^');
        }

        if ((Modifiers & KeyModifiers.Alt) != 0)
        {
            sb.Append('%');
        }

        if ((Modifiers & KeyModifiers.Shift) != 0)
        {
            sb.Append('+');
        }

        if (NamedKeys.TryGetValue(Key, out var code))
        {
            sb.Append(code);
        }
        else if (BracedInOnKey.IndexOf(Key[0]) >= 0)
        {
            sb.Append('{').Append(Key).Append('}');
        }
        else
        {
            sb.Append(Key.ToLowerInvariant());
        }

        return sb.ToString();
    }

    /// <summary>
    /// Renders the chord for people, with modifiers in the order Ctrl, Alt, Shift:
    /// e.g. <c>Ctrl+Shift+K</c>, <c>Ctrl+Alt+Shift+Up</c>. <see cref="Parse"/> accepts this form back.
    /// </summary>
    public string ToDisplayString()
    {
        var sb = new StringBuilder(24);
        if ((Modifiers & KeyModifiers.Ctrl) != 0)
        {
            sb.Append("Ctrl+");
        }

        if ((Modifiers & KeyModifiers.Alt) != 0)
        {
            sb.Append("Alt+");
        }

        if ((Modifiers & KeyModifiers.Shift) != 0)
        {
            sb.Append("Shift+");
        }

        return sb.Append(Key).ToString();
    }

    /// <summary>
    /// The keys this chord is pressed with on a US keyboard: a shifted punctuation key becomes Shift plus the key's
    /// unshifted character (<c>Ctrl+{</c> and <c>Ctrl+Shift+{</c> give <c>Ctrl+Shift+[</c>, <c>Ctrl++</c> gives
    /// <c>Ctrl+Shift+=</c>); any other chord is returned unchanged. Two chords with equal results are the same
    /// key press there, so they must not be bound to different actions.
    /// </summary>
    public KeyChord ToUsKeys()
    {
        var index = Key.Length == 1 ? UsShifted.IndexOf(Key[0]) : -1;
        return index < 0 ? this : new KeyChord(Modifiers | KeyModifiers.Shift, UsUnshifted[index].ToString());
    }

    /// <summary>Same as <see cref="ToDisplayString"/>.</summary>
    public override string ToString() => ToDisplayString();

    /// <inheritdoc />
    public bool Equals(KeyChord? other) =>
        other is not null && Modifiers == other.Modifiers && string.Equals(Key, other.Key, StringComparison.Ordinal);

    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as KeyChord);

    /// <inheritdoc />
    public override int GetHashCode() => ((int)Modifiers * 397) ^ StringComparer.Ordinal.GetHashCode(Key);

    /// <summary>Value equality.</summary>
    public static bool operator ==(KeyChord? left, KeyChord? right) => left is null ? right is null : left.Equals(right);

    /// <summary>Value inequality.</summary>
    public static bool operator !=(KeyChord? left, KeyChord? right) => !(left == right);

    private static string ParseKey(string text, string token)
    {
        if (token.Length == 0)
        {
            throw Invalid(text, "it has no key after the last '+'.");
        }

        if (token.Length == 1)
        {
            var c = token[0];
            if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
            {
                return char.ToUpperInvariant(c).ToString();
            }

            if ((c >= '0' && c <= '9') || Punctuation.IndexOf(c) >= 0)
            {
                return token;
            }
        }

        if (NamedKeyAliases.TryGetValue(token, out var named))
        {
            return named;
        }

        if (TryParseModifier(token, out _))
        {
            throw Invalid(text, $"'{token}' is a modifier; a key must follow the modifiers.");
        }

        throw Invalid(
            text,
            $"'{token}' is not a known key. Use a letter, digit, punctuation character, " +
            "Up, Down, Left, Right, PgUp, PgDn, Home, End, Ins, Del or F1-F12.");
    }

    private static KeyModifiers ParseModifiers(string text, string? modifierPart)
    {
        if (modifierPart is null)
        {
            throw Invalid(text, "it has no modifier. Use Ctrl, Alt and/or Shift, e.g. \"Ctrl+Shift+K\".");
        }

        var result = KeyModifiers.None;
        foreach (var raw in modifierPart.Split('+'))
        {
            var token = raw.Trim();
            if (token.Length == 0)
            {
                throw Invalid(text, "it has an empty modifier (two '+' in a row).");
            }

            if (!TryParseModifier(token, out var modifier))
            {
                throw Invalid(text, $"'{token}' is not a modifier. Use Ctrl, Alt or Shift.");
            }

            if ((result & modifier) != 0)
            {
                throw Invalid(text, $"the modifier '{modifier}' appears more than once.");
            }

            result |= modifier;
        }

        return result;
    }

    private static bool TryParseModifier(string token, out KeyModifiers modifier)
    {
        switch (token.ToUpperInvariant())
        {
            case "CTRL":
            case "CONTROL":
                modifier = KeyModifiers.Ctrl;
                return true;
            case "ALT":
                modifier = KeyModifiers.Alt;
                return true;
            case "SHIFT":
                modifier = KeyModifiers.Shift;
                return true;
            default:
                modifier = KeyModifiers.None;
                return false;
        }
    }

    private static FormatException Invalid(string text, string reason) =>
        new FormatException($"Invalid key chord \"{text}\": {reason}");

    private static Dictionary<string, string> CreateNamedKeys()
    {
        var keys = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["Up"] = "{UP}",
            ["Down"] = "{DOWN}",
            ["Left"] = "{LEFT}",
            ["Right"] = "{RIGHT}",
            ["PgUp"] = "{PGUP}",
            ["PgDn"] = "{PGDN}",
            ["Home"] = "{HOME}",
            ["End"] = "{END}",
            ["Ins"] = "{INSERT}",
            ["Del"] = "{DELETE}",
        };
        for (var i = 1; i <= 12; i++)
        {
            var name = "F" + i.ToString(CultureInfo.InvariantCulture);
            keys[name] = "{" + name + "}";
        }

        return keys;
    }

    private static Dictionary<string, string> CreateNamedKeyAliases()
    {
        var aliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in NamedKeys.Keys)
        {
            aliases[name] = name;
        }

        aliases["PageUp"] = "PgUp";
        aliases["PageDown"] = "PgDn";
        aliases["PgDown"] = "PgDn";
        aliases["Insert"] = "Ins";
        aliases["Delete"] = "Del";
        return aliases;
    }
}
