using ExcelModelingToolkit.Core.Keys;
using Xunit;

namespace ExcelModelingToolkit.Core.Tests.Keys;

public class KeyCaptureTests
{
    /// <summary>Windows virtual-key code and the key-string name it captures as.</summary>
    public static TheoryData<int, string> NamedKeys => new TheoryData<int, string>
    {
        { 0x41, "A" },
        { 0x4B, "K" },
        { 0x5A, "Z" },
        { 0x30, "0" },
        { 0x31, "1" },
        { 0x39, "9" },
        { 0x70, "F1" },
        { 0x7B, "F12" },
        { 0x7C, "F13" },
        { 0x87, "F24" },
        { 0x21, "PgUp" },
        { 0x22, "PgDn" },
        { 0x23, "End" },
        { 0x24, "Home" },
        { 0x25, "Left" },
        { 0x26, "Up" },
        { 0x27, "Right" },
        { 0x28, "Down" },
        { 0x2D, "Ins" },
        { 0x2E, "Del" },
        { 0xBA, ";" },  // VK_OEM_1
        { 0xBB, "=" },  // VK_OEM_PLUS
        { 0xBC, "," },  // VK_OEM_COMMA
        { 0xBD, "-" },  // VK_OEM_MINUS
        { 0xBE, "." },  // VK_OEM_PERIOD
        { 0xBF, "/" },  // VK_OEM_2
        { 0xC0, "`" },  // VK_OEM_3
        { 0xDB, "[" },  // VK_OEM_4
        { 0xDC, "\\" }, // VK_OEM_5
        { 0xDD, "]" },  // VK_OEM_6
        { 0xDE, "'" },  // VK_OEM_7
    };

    [Theory]
    [MemberData(nameof(NamedKeys))]
    public void Maps_virtual_keys_to_key_names(int virtualKey, string name)
    {
        Assert.Equal(name, KeyCapture.KeyName(virtualKey));
        Assert.Equal("Ctrl+" + name, KeyCapture.ToKeyString(virtualKey, KeyModifiers.Ctrl));
    }

    [Theory]
    [InlineData(0x10)] // VK_SHIFT
    [InlineData(0x11)] // VK_CONTROL
    [InlineData(0x12)] // VK_MENU (Alt)
    [InlineData(0xA0)] // VK_LSHIFT
    [InlineData(0xA1)] // VK_RSHIFT
    [InlineData(0xA2)] // VK_LCONTROL
    [InlineData(0xA3)] // VK_RCONTROL
    [InlineData(0xA4)] // VK_LMENU
    [InlineData(0xA5)] // VK_RMENU
    [InlineData(0x5B)] // VK_LWIN
    [InlineData(0x5C)] // VK_RWIN
    public void Modifier_only_presses_give_null(int virtualKey)
    {
        Assert.True(KeyCapture.IsModifierKey(virtualKey));
        Assert.Null(KeyCapture.KeyName(virtualKey));
        Assert.Null(KeyCapture.ToKeyString(virtualKey, KeyModifiers.Ctrl | KeyModifiers.Shift));
    }

    [Theory]
    [InlineData(0x0D)] // Enter
    [InlineData(0x09)] // Tab
    [InlineData(0x20)] // Space
    [InlineData(0x1B)] // Esc
    [InlineData(0x08)] // Backspace
    [InlineData(0x60)] // Numpad 0
    [InlineData(0x6B)] // Numpad +
    [InlineData(0xE2)] // VK_OEM_102
    [InlineData(0x14)] // Caps Lock
    public void Keys_without_a_name_give_null(int virtualKey)
    {
        Assert.False(KeyCapture.IsModifierKey(virtualKey));
        Assert.Null(KeyCapture.KeyName(virtualKey));
        Assert.Null(KeyCapture.ToKeyString(virtualKey, KeyModifiers.Ctrl));
    }

    [Theory]
    [InlineData(0xDB, KeyModifiers.Ctrl | KeyModifiers.Shift, "Ctrl+Shift+[", "^+{[}")]
    [InlineData(0xDE, KeyModifiers.Ctrl, "Ctrl+'", "^'")]
    [InlineData(0xBA, KeyModifiers.Ctrl, "Ctrl+;", "^;")]
    [InlineData(0x31, KeyModifiers.Ctrl | KeyModifiers.Shift, "Ctrl+Shift+1", "^+1")]
    [InlineData(0x4B, KeyModifiers.Shift | KeyModifiers.Ctrl, "Ctrl+Shift+K", "^+k")]
    [InlineData(0x7B, KeyModifiers.Shift | KeyModifiers.Alt | KeyModifiers.Ctrl, "Ctrl+Alt+Shift+F12", "^%+{F12}")]
    [InlineData(0xDC, KeyModifiers.Alt, "Alt+\\", "%\\")]
    [InlineData(0x21, KeyModifiers.Shift, "Shift+PgUp", "+{PGUP}")]
    public void Captured_chords_parse_to_the_expected_binding(int virtualKey, KeyModifiers modifiers, string expected, string onKey)
    {
        var text = KeyCapture.ToKeyString(virtualKey, modifiers);

        Assert.Equal(expected, text);
        var chord = KeyChord.Parse(text!);
        Assert.Equal(expected, chord.ToDisplayString());
        Assert.Equal(onKey, chord.ToOnKeyString());
    }

    [Fact]
    public void Unbindable_presses_are_captured_as_typed_and_rejected_by_the_parser()
    {
        // Capture does not validate; KeyChord.Parse gives the user-facing reason.
        Assert.Equal("K", KeyCapture.ToKeyString(0x4B, KeyModifiers.None));
        Assert.Equal("Shift+K", KeyCapture.ToKeyString(0x4B, KeyModifiers.Shift));
        Assert.Equal("Ctrl+F13", KeyCapture.ToKeyString(0x7C, KeyModifiers.Ctrl));

        Assert.Throws<System.FormatException>(() => KeyChord.Parse("K"));
        Assert.Throws<System.FormatException>(() => KeyChord.Parse("Shift+K"));
        Assert.Throws<System.FormatException>(() => KeyChord.Parse("Ctrl+F13"));
    }
}
