using System;
using Modelwright.Core.Keys;
using Xunit;

namespace Modelwright.Core.Tests.Keys;

public class KeyChordTests
{
    /// <summary>
    /// Every binding the Excel-DNA spike registered successfully through xlcOnKey (spikes/exceldna/AddIn.cs,
    /// Keys and Coverage tables; docs/spike-results.md K3), as (human key, proven OnKey string). The spike's
    /// upper-case <c>^+K</c> hedge is intentionally excluded: KeyChord always emits lower-case letters (<c>^+k</c>).
    /// </summary>
    public static TheoryData<string, string> SpikeProvenKeys => new TheoryData<string, string>
    {
        // Keys table: Macabacus v1 keys and hedges
        { "Ctrl+Shift+1", "^+1" },
        { "Ctrl+Shift+2", "^+2" },
        { "Ctrl+Shift+4", "^+4" },
        { "Ctrl+Shift+5", "^+5" },
        { "Ctrl+Shift+8", "^+8" },
        { "Ctrl+'", "^'" },
        { "Ctrl+Shift+K", "^+k" },
        { "Ctrl+;", "^;" },
        { "Ctrl+,", "^," },
        { "Ctrl+.", "^." },
        { "Ctrl+Shift+[", "^+{[}" },
        { "Ctrl+{", "^{{}" },
        { "Ctrl+Alt+[", "^%{[}" },
        { "Ctrl+Shift+\\", "^+\\" },
        { "Ctrl+|", "^|" },

        // Keys table: spike-only commands on Ctrl+Alt+Shift+F-keys
        { "Ctrl+Alt+Shift+F1", "^%+{F1}" },
        { "Ctrl+Alt+Shift+F2", "^%+{F2}" },
        { "Ctrl+Alt+Shift+F3", "^%+{F3}" },
        { "Ctrl+Alt+Shift+F5", "^%+{F5}" },
        { "Ctrl+Alt+Shift+F6", "^%+{F6}" },
        { "Ctrl+Alt+Shift+F7", "^%+{F7}" },
        { "Ctrl+Alt+Shift+F8", "^%+{F8}" },
        { "Ctrl+Alt+Shift+F9", "^%+{F9}" },
        { "Ctrl+Alt+Shift+F10", "^%+{F10}" },
        { "Ctrl+Alt+Shift+F11", "^%+{F11}" },
        { "Ctrl+Alt+Shift+F12", "^%+{F12}" },

        // Coverage table: remaining Macabacus punctuation and named keys
        { "Alt+Shift+;", "%+;" }, // Ratio cycle (Macabacus factory key)
        { "Alt+Shift+,", "%+," },
        { "Alt+Shift+.", "%+." },
        { "Ctrl+Alt+Shift+'", "^%+'" },
        { "Ctrl+Alt+.", "^%." },
        { "Ctrl+Alt+'", "^%'" },
        { "Ctrl+Shift+]", "^+{]}" },
        { "Ctrl+Alt+Shift+[", "^%+{[}" },
        { "Ctrl+Alt+Shift+]", "^%+{]}" },
        { "Ctrl+Alt+]", "^%{]}" },
        { "Ctrl+Alt+\\", "^%\\" },
        { "Ctrl+Alt+=", "^%=" },
        { "Ctrl+Alt+-", "^%-" },
        { "Ctrl+Alt+Shift+,", "^%+," },
        { "Ctrl+Alt+Shift+.", "^%+." },
        { "Alt+Shift+=", "%+=" },
        { "Ctrl+Alt+Shift+=", "^%+=" },
        { "Alt+Shift+-", "%+-" },
        { "Ctrl+Alt+Shift+-", "^%+-" },
        { "Ctrl+Alt+Shift+Up", "^%+{UP}" },
        { "Alt+Shift+PgUp", "%+{PGUP}" },
        { "Ctrl+Alt+Shift+Ins", "^%+{INSERT}" },
        { "Ctrl+F2", "^{F2}" },
        { "Alt+F12", "%{F12}" },
        { "Ctrl+Alt+Home", "^%{HOME}" },
        { "Ctrl+Shift+Y", "^+y" }, // Binary cycle (Macabacus factory key)
    };

    /// <summary>The v1 keymap of docs/PLAN.md section 4.2, including undo/redo and the v2 auditing keys.</summary>
    public static TheoryData<string, string> PlanV1Keymap => new TheoryData<string, string>
    {
        { "Ctrl+Shift+1", "^+1" }, // General Number cycle
        { "Ctrl+Shift+2", "^+2" }, // Date cycle
        { "Ctrl+Shift+4", "^+4" }, // Local Currency cycle
        { "Ctrl+Shift+5", "^+5" }, // Percent cycle
        { "Ctrl+Shift+8", "^+8" }, // Multiple cycle
        { "Ctrl+'", "^'" }, // Font Color cycle
        { "Ctrl+Shift+K", "^+k" }, // Fill Color cycle
        { "Ctrl+Shift+[", "^+{[}" }, // Pro Precedents (Trace In)
        { "Ctrl+{", "^{{}" }, // Pro Precedents hedge
        { "Ctrl+Shift+\\", "^+\\" }, // Last Audited Cell
        { "Ctrl+|", "^|" }, // Last Audited Cell hedge
        { "Ctrl+;", "^;" }, // Blue-Black toggle
        { "Ctrl+,", "^," }, // Increase decimals
        { "Ctrl+.", "^." }, // Decrease decimals
        { "Ctrl+Z", "^z" }, // Undo (formatting stack)
        { "Ctrl+Y", "^y" }, // Redo
        { "Ctrl+Shift+]", "^+{]}" }, // Pro Dependents (v2)
        { "Ctrl+Alt+[", "^%{[}" }, // Show All Precedents (v2)
        { "Ctrl+Alt+]", "^%{]}" }, // Show All Dependents (v2)
    };

    /// <summary>
    /// Further keys from the full Macabacus keymap (docs/research/06) that exercise every named key and
    /// the remaining punctuation, so they stay bindable when later features claim them.
    /// </summary>
    public static TheoryData<string, string> MacabacusKeymapExtras => new TheoryData<string, string>
    {
        { "Ctrl+Alt+Shift+Down", "^%+{DOWN}" }, // Bottom Border
        { "Ctrl+Alt+Shift+Left", "^%+{LEFT}" }, // Left Border
        { "Ctrl+Alt+Shift+Right", "^%+{RIGHT}" }, // Right Border
        { "Ctrl+Alt+Left", "^%{LEFT}" }, // Quick Export To PowerPoint
        { "Ctrl+Alt+Right", "^%{RIGHT}" }, // Quick Export To Word
        { "Ctrl+Alt+End", "^%{END}" }, // Last Sheet
        { "Ctrl+Alt+PgDn", "^%{PGDN}" }, // Next Sheet Loop
        { "Ctrl+Alt+PgUp", "^%{PGUP}" }, // Previous Sheet Loop
        { "Ctrl+Alt+Shift+PgUp", "^%+{PGUP}" }, // Column Width Cycle
        { "Alt+Shift+PgDn", "%+{PGDN}" }, // AutoFit Height
        { "Ctrl+Alt+Shift+PgDn", "^%+{PGDN}" }, // AutoFit Width
        { "Alt+Shift+Ins", "%+{INSERT}" }, // Insert Row
        { "Alt+Shift+Del", "%+{DELETE}" }, // Delete Row
        { "Ctrl+Alt+Shift+Del", "^%+{DELETE}" }, // Delete Column
        { "Alt+Shift+Home", "%+{HOME}" }, // Hide Row
        { "Ctrl+Alt+Shift+Home", "^%+{HOME}" }, // Hide Column
        { "Alt+Shift+End", "%+{END}" }, // Unhide Row
        { "Ctrl+Alt+Shift+End", "^%+{END}" }, // Unhide Column
        { "Shift+F12", "+{F12}" }, // Quick Save Up
        { "Ctrl+Shift+-", "^+-" }, // No Border
        { "Ctrl+Shift+7", "^+7" }, // Outside Borders
        { "Ctrl+Alt+Shift+7", "^%+7" }, // Inside Borders
        { "Ctrl+Alt+9", "^%9" }, // Wrap Parentheses
        { "Ctrl+Q", "^q" }, // Uniformulas
        { "Ctrl+Alt+C", "^%c" }, // Chart Color Cycle
        { "Alt+Shift+G", "%+g" }, // Font Size Cycle
    };

    [Theory]
    [MemberData(nameof(SpikeProvenKeys))]
    public void Spike_proven_keys_render_and_round_trip(string human, string onKey) => AssertChord(human, onKey);

    [Theory]
    [MemberData(nameof(PlanV1Keymap))]
    public void Plan_keymap_renders_and_round_trips(string human, string onKey) => AssertChord(human, onKey);

    [Theory]
    [MemberData(nameof(MacabacusKeymapExtras))]
    public void Macabacus_keymap_extras_render_and_round_trip(string human, string onKey) => AssertChord(human, onKey);

    [Theory]
    [InlineData("ctrl+shift+k", "Ctrl+Shift+K", "^+k")]
    [InlineData("CTRL+SHIFT+K", "Ctrl+Shift+K", "^+k")]
    [InlineData("Shift+Ctrl+K", "Ctrl+Shift+K", "^+k")]
    [InlineData("Shift+Alt+Ctrl+Up", "Ctrl+Alt+Shift+Up", "^%+{UP}")]
    [InlineData("Control+Alt+Shift+up", "Ctrl+Alt+Shift+Up", "^%+{UP}")]
    [InlineData("  Ctrl + Shift + [  ", "Ctrl+Shift+[", "^+{[}")]
    [InlineData("Alt+Shift+PageUp", "Alt+Shift+PgUp", "%+{PGUP}")]
    [InlineData("Alt+Shift+pgup", "Alt+Shift+PgUp", "%+{PGUP}")]
    [InlineData("Alt+Shift+PageDown", "Alt+Shift+PgDn", "%+{PGDN}")]
    [InlineData("Alt+Shift+PgDown", "Alt+Shift+PgDn", "%+{PGDN}")]
    [InlineData("Ctrl+Alt+Shift+Insert", "Ctrl+Alt+Shift+Ins", "^%+{INSERT}")]
    [InlineData("Alt+Shift+Delete", "Alt+Shift+Del", "%+{DELETE}")]
    [InlineData("ctrl+f2", "Ctrl+F2", "^{F2}")]
    [InlineData("alt+F12", "Alt+F12", "%{F12}")]
    public void Normalizes_case_order_spacing_and_aliases(string input, string display, string onKey)
    {
        var chord = KeyChord.Parse(input);

        Assert.Equal(display, chord.ToDisplayString());
        Assert.Equal(onKey, chord.ToOnKeyString());
        Assert.Equal(KeyChord.Parse(display), chord);
    }

    [Theory]
    [InlineData("Ctrl++", "^{+}")]
    [InlineData("Ctrl+Alt++", "^%{+}")]
    [InlineData("Ctrl + +", "^{+}")]
    [InlineData("Ctrl+^", "^{^}")]
    [InlineData("Ctrl+%", "^{%}")]
    [InlineData("Ctrl+~", "^{~}")]
    [InlineData("Ctrl+(", "^{(}")]
    [InlineData("Ctrl+)", "^{)}")]
    [InlineData("Ctrl+{", "^{{}")]
    [InlineData("Ctrl+}", "^{}}")]
    [InlineData("Ctrl+[", "^{[}")]
    [InlineData("Ctrl+]", "^{]}")]
    public void Braces_onkey_syntax_characters(string human, string onKey)
    {
        var chord = KeyChord.Parse(human);

        Assert.Equal(onKey, chord.ToOnKeyString());
        Assert.Equal(chord, KeyChord.Parse(chord.ToDisplayString()));
    }

    [Theory]
    [InlineData("Ctrl+`", "^`")]
    [InlineData("Ctrl+/", "^/")]
    [InlineData("Ctrl+=", "^=")]
    [InlineData("Ctrl+-", "^-")]
    [InlineData("Ctrl+!", "^!")]
    [InlineData("Ctrl+@", "^@")]
    [InlineData("Ctrl+#", "^#")]
    [InlineData("Ctrl+$", "^$")]
    [InlineData("Ctrl+&", "^&")]
    [InlineData("Ctrl+*", "^*")]
    [InlineData("Ctrl+_", "^_")]
    [InlineData("Ctrl+:", "^:")]
    [InlineData("Ctrl+\"", "^\"")]
    [InlineData("Ctrl+<", "^<")]
    [InlineData("Ctrl+>", "^>")]
    [InlineData("Ctrl+?", "^?")]
    [InlineData("Ctrl+0", "^0")]
    [InlineData("Ctrl+9", "^9")]
    [InlineData("Ctrl+a", "^a")]
    [InlineData("Ctrl+Z", "^z")]
    public void Leaves_other_characters_unbraced(string human, string onKey)
    {
        Assert.Equal(onKey, KeyChord.Parse(human).ToOnKeyString());
    }

    [Fact]
    public void Renders_all_function_keys()
    {
        for (var i = 1; i <= 12; i++)
        {
            var chord = KeyChord.Parse($"Ctrl+F{i}");

            Assert.Equal($"^{{F{i}}}", chord.ToOnKeyString());
            Assert.Equal($"Ctrl+F{i}", chord.ToDisplayString());
            Assert.True(chord.IsNamedKey);
        }
    }

    [Fact]
    public void Exposes_modifiers_and_key()
    {
        var chord = KeyChord.Parse("Ctrl+Alt+Shift+Up");

        Assert.Equal(KeyModifiers.Ctrl | KeyModifiers.Alt | KeyModifiers.Shift, chord.Modifiers);
        Assert.Equal("Up", chord.Key);
        Assert.True(chord.IsNamedKey);

        var letter = KeyChord.Parse("Ctrl+Shift+k");
        Assert.Equal(KeyModifiers.Ctrl | KeyModifiers.Shift, letter.Modifiers);
        Assert.Equal("K", letter.Key);
        Assert.False(letter.IsNamedKey);

        Assert.Equal(KeyModifiers.Alt, KeyChord.Parse("Alt+F12").Modifiers);
    }

    [Fact]
    public void Shift_alone_is_allowed_with_named_keys()
    {
        var chord = KeyChord.Parse("Shift+F12");

        Assert.Equal(KeyModifiers.Shift, chord.Modifiers);
        Assert.Equal("+{F12}", chord.ToOnKeyString());
        Assert.Equal("+{UP}", KeyChord.Parse("Shift+Up").ToOnKeyString());
    }

    [Fact]
    public void Distinguishes_shifted_character_from_shift_modifier()
    {
        var shiftBackslash = KeyChord.Parse("Ctrl+Shift+\\");
        var pipe = KeyChord.Parse("Ctrl+|");

        Assert.NotEqual(shiftBackslash, pipe);
        Assert.Equal("^+\\", shiftBackslash.ToOnKeyString());
        Assert.Equal("^|", pipe.ToOnKeyString());
    }

    [Theory]
    [InlineData("K")]
    [InlineData("[")]
    [InlineData("F2")]
    [InlineData("Up")]
    [InlineData("+")]
    [InlineData("'")]
    public void Rejects_key_without_modifier(string text)
    {
        var ex = Assert.Throws<FormatException>(() => KeyChord.Parse(text));
        Assert.Contains("no modifier", ex.Message);
        Assert.Contains($"\"{text}\"", ex.Message);
    }

    [Theory]
    [InlineData("Ctrl+Ctrl+K", "Ctrl")]
    [InlineData("Ctrl+Control+K", "Ctrl")]
    [InlineData("Alt+Shift+alt+;", "Alt")]
    [InlineData("Shift+Ctrl+SHIFT+[", "Shift")]
    public void Rejects_duplicate_modifier(string text, string modifier)
    {
        var ex = Assert.Throws<FormatException>(() => KeyChord.Parse(text));
        Assert.Contains($"'{modifier}' appears more than once", ex.Message);
    }

    [Theory]
    [InlineData("Ctrl+Foo")]
    [InlineData("Ctrl+F0")]
    [InlineData("Ctrl+F13")]
    [InlineData("Ctrl+Esc")]
    [InlineData("Ctrl+Space")]
    [InlineData("Ctrl+ab")]
    [InlineData("Ctrl+é")]
    [InlineData("Ctrl+£")]
    [InlineData("Ctrl+Shift+PgUpp")]
    public void Rejects_unknown_key(string text)
    {
        var ex = Assert.Throws<FormatException>(() => KeyChord.Parse(text));
        Assert.Contains("is not a known key", ex.Message);
    }

    [Theory]
    [InlineData("Win+K")]
    [InlineData("Cmd+K")]
    [InlineData("Ctl+K")]
    [InlineData("Ctrl+Meta+K")]
    public void Rejects_unknown_modifier(string text)
    {
        var ex = Assert.Throws<FormatException>(() => KeyChord.Parse(text));
        Assert.Contains("is not a modifier", ex.Message);
    }

    [Theory]
    [InlineData("Ctrl+Shift")]
    [InlineData("Ctrl+Alt")]
    [InlineData("Shift+Control")]
    public void Rejects_modifier_as_key(string text)
    {
        var ex = Assert.Throws<FormatException>(() => KeyChord.Parse(text));
        Assert.Contains("is a modifier", ex.Message);
    }

    [Theory]
    [InlineData("Ctrl+")]
    [InlineData("Ctrl+Shift+")]
    [InlineData("Ctrl+ ")]
    public void Rejects_missing_key(string text)
    {
        var ex = Assert.Throws<FormatException>(() => KeyChord.Parse(text));
        Assert.Contains("no key", ex.Message);
    }

    [Theory]
    [InlineData("Ctrl++K")]
    [InlineData("+Ctrl+K")]
    [InlineData("Ctrl+ +K")]
    public void Rejects_empty_modifier(string text)
    {
        var ex = Assert.Throws<FormatException>(() => KeyChord.Parse(text));
        Assert.Contains("empty modifier", ex.Message);
    }

    [Theory]
    [InlineData("Shift+K")]
    [InlineData("Shift+1")]
    [InlineData("Shift+[")]
    [InlineData("Shift++")]
    public void Rejects_shift_alone_with_character_key(string text)
    {
        var ex = Assert.Throws<FormatException>(() => KeyChord.Parse(text));
        Assert.Contains("block typing", ex.Message);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Rejects_empty_text(string text)
    {
        var ex = Assert.Throws<FormatException>(() => KeyChord.Parse(text));
        Assert.Contains("empty", ex.Message);
    }

    [Fact]
    public void Rejects_null()
    {
        Assert.Throws<ArgumentNullException>(() => KeyChord.Parse(null!));
    }

    [Fact]
    public void Equality_is_by_value()
    {
        var a = KeyChord.Parse("Ctrl+Shift+K");
        var b = KeyChord.Parse("shift+ctrl+k");
        var c = KeyChord.Parse("Ctrl+K");
        KeyChord? none = null;

        Assert.True(a == b);
        Assert.False(a != b);
        Assert.True(a.Equals((object)b));
        Assert.Equal(a.GetHashCode(), b.GetHashCode());
        Assert.True(a != c);
        Assert.False(a.Equals(c));
        Assert.False(a.Equals(null));
        Assert.False(a.Equals("Ctrl+Shift+K"));
        Assert.False(a == none);
        Assert.False(none == a);
        Assert.True(none == null);
    }

    [Theory]
    [InlineData("Ctrl+{", "Ctrl+Shift+[")]
    [InlineData("Ctrl+Shift+{", "Ctrl+Shift+[")]
    [InlineData("Ctrl++", "Ctrl+Shift+=")]
    [InlineData("Ctrl+|", "Ctrl+Shift+\\")]
    [InlineData("Alt+~", "Alt+Shift+`")]
    [InlineData("Ctrl+!", "Ctrl+Shift+1")]
    [InlineData("Ctrl+)", "Ctrl+Shift+0")]
    [InlineData("Ctrl+_", "Ctrl+Shift+-")]
    [InlineData("Ctrl+:", "Ctrl+Shift+;")]
    [InlineData("Ctrl+\"", "Ctrl+Shift+'")]
    [InlineData("Ctrl+<", "Ctrl+Shift+,")]
    [InlineData("Ctrl+>", "Ctrl+Shift+.")]
    [InlineData("Ctrl+?", "Ctrl+Shift+/")]
    [InlineData("Ctrl+Alt+^", "Ctrl+Alt+Shift+6")]
    public void Shifted_punctuation_is_the_shifted_key_on_a_us_keyboard(string chord, string usKeys)
    {
        Assert.Equal(usKeys, KeyChord.Parse(chord).ToUsKeys().ToDisplayString());
        Assert.Equal(KeyChord.Parse(usKeys), KeyChord.Parse(chord).ToUsKeys());
        Assert.NotEqual(KeyChord.Parse(usKeys), KeyChord.Parse(chord)); // still different chords to OnKey
    }

    [Theory]
    [InlineData("Ctrl+Shift+[")]
    [InlineData("Ctrl+[")]
    [InlineData("Ctrl+K")]
    [InlineData("Ctrl+Shift+1")]
    [InlineData("Alt+PgUp")]
    [InlineData("Ctrl+F2")]
    public void Unshifted_keys_are_their_own_us_keys(string chord)
    {
        var parsed = KeyChord.Parse(chord);

        Assert.Same(parsed, parsed.ToUsKeys());
    }

    /// <summary>Parses <paramref name="human"/>, checks the OnKey string, and round-trips the display string.</summary>
    private static void AssertChord(string human, string onKey)
    {
        var chord = KeyChord.Parse(human);

        Assert.Equal(onKey, chord.ToOnKeyString());
        Assert.Equal(human, chord.ToDisplayString());
        Assert.Equal(human, chord.ToString());

        var reparsed = KeyChord.Parse(chord.ToDisplayString());
        Assert.Equal(chord, reparsed);
        Assert.Equal(onKey, reparsed.ToOnKeyString());
    }
}
