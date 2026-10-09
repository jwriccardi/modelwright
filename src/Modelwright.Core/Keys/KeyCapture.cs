using System.Globalization;
using System.Text;

namespace Modelwright.Core.Keys;

/// <summary>
/// Turns a key press captured from the keyboard (a Windows virtual-key code plus the modifier state) into the key
/// string syntax that <see cref="KeyChord.Parse"/> reads, e.g. <c>Ctrl+Shift+[</c>, so a shortcut can be recorded
/// by pressing it. The result is not validated: <see cref="KeyChord.Parse"/> still decides whether it can be bound
/// (it rejects a key with no modifier, Shift alone with a character key, and F13-F24, which Excel's
/// <c>OnKey</c> cannot bind).
/// </summary>
/// <remarks>
/// <para>
/// Punctuation keys are named by their <b>US keyboard layout</b> character, which is how Excel's <c>OnKey</c> and
/// our key strings name them: <c>VK_OEM_1</c> <c>;</c>, <c>VK_OEM_PLUS</c> <c>=</c>, <c>VK_OEM_COMMA</c> <c>,</c>,
/// <c>VK_OEM_MINUS</c> <c>-</c>, <c>VK_OEM_PERIOD</c> <c>.</c>, <c>VK_OEM_2</c> <c>/</c>, <c>VK_OEM_3</c> <c>`</c>,
/// <c>VK_OEM_4</c> <c>[</c>, <c>VK_OEM_5</c> <c>\</c>, <c>VK_OEM_6</c> <c>]</c>, <c>VK_OEM_7</c> <c>'</c>. Other
/// layouts assign these virtual-key codes to keys of their own choosing, so there the captured name follows the
/// virtual-key code, not the character printed on the key. This class stays US-only (it has no access to the
/// active keyboard layout); a caller that has, such as the add-in's Capture button, substitutes the layout's
/// character for a punctuation key and warns about Ctrl+Alt (AltGr) combinations that type a character.
/// </para>
/// <para>
/// The unshifted character is always used, with Shift as a modifier: Ctrl+Shift+[ is <c>Ctrl+Shift+[</c>, never
/// <c>Ctrl+{</c>, matching the Macabacus keymap.
/// </para>
/// <para>
/// Also mapped: letters, the top-row digits, F1-F24, the arrows, PgUp, PgDn, Home, End, Ins and Del. Other keys
/// (Enter, Tab, Space, Esc, Backspace, the numeric keypad, the Windows keys) have no key-string name and give null,
/// as does a press of a modifier key alone.
/// </para>
/// </remarks>
public static class KeyCapture
{
    private const int VkShift = 0x10;
    private const int VkControl = 0x11;
    private const int VkMenu = 0x12; // Alt
    private const int VkPrior = 0x21; // PgUp
    private const int VkNext = 0x22; // PgDn
    private const int VkEnd = 0x23;
    private const int VkHome = 0x24;
    private const int VkLeft = 0x25;
    private const int VkUp = 0x26;
    private const int VkRight = 0x27;
    private const int VkDown = 0x28;
    private const int VkInsert = 0x2D;
    private const int VkDelete = 0x2E;
    private const int VkLWin = 0x5B;
    private const int VkRWin = 0x5C;
    private const int VkF1 = 0x70;
    private const int VkF24 = 0x87;
    private const int VkLShift = 0xA0;
    private const int VkRMenu = 0xA5; // VK_LSHIFT..VK_RMENU are the left/right Shift, Ctrl and Alt keys.
    private const int VkOem1 = 0xBA;
    private const int VkOemPlus = 0xBB;
    private const int VkOemComma = 0xBC;
    private const int VkOemMinus = 0xBD;
    private const int VkOemPeriod = 0xBE;
    private const int VkOem2 = 0xBF;
    private const int VkOem3 = 0xC0;
    private const int VkOem4 = 0xDB;
    private const int VkOem5 = 0xDC;
    private const int VkOem6 = 0xDD;
    private const int VkOem7 = 0xDE;

    /// <summary>True for Shift, Ctrl, Alt (either side) and the Windows keys: pressed alone, they are not a shortcut yet.</summary>
    public static bool IsModifierKey(int virtualKey) =>
        virtualKey == VkShift || virtualKey == VkControl || virtualKey == VkMenu ||
        (virtualKey >= VkLShift && virtualKey <= VkRMenu) ||
        virtualKey == VkLWin || virtualKey == VkRWin;

    /// <summary>
    /// The key-string name of <paramref name="virtualKey"/> (e.g. <c>K</c>, <c>1</c>, <c>[</c>, <c>PgUp</c>,
    /// <c>F2</c>), or null for a modifier or a key with no name (see the remarks).
    /// </summary>
    public static string? KeyName(int virtualKey)
    {
        if ((virtualKey >= 'A' && virtualKey <= 'Z') || (virtualKey >= '0' && virtualKey <= '9'))
        {
            return ((char)virtualKey).ToString();
        }

        if (virtualKey >= VkF1 && virtualKey <= VkF24)
        {
            return "F" + (virtualKey - VkF1 + 1).ToString(CultureInfo.InvariantCulture);
        }

        return virtualKey switch
        {
            VkPrior => "PgUp",
            VkNext => "PgDn",
            VkEnd => "End",
            VkHome => "Home",
            VkLeft => "Left",
            VkUp => "Up",
            VkRight => "Right",
            VkDown => "Down",
            VkInsert => "Ins",
            VkDelete => "Del",
            VkOem1 => ";",
            VkOemPlus => "=",
            VkOemComma => ",",
            VkOemMinus => "-",
            VkOemPeriod => ".",
            VkOem2 => "/",
            VkOem3 => "`",
            VkOem4 => "[",
            VkOem5 => "\\",
            VkOem6 => "]",
            VkOem7 => "'",
            _ => null,
        };
    }

    /// <summary>
    /// The key string for a press of <paramref name="virtualKey"/> with <paramref name="modifiers"/> held, with the
    /// modifiers in the order Ctrl, Alt, Shift (e.g. <c>Ctrl+Shift+[</c>), or null if the key is a modifier
    /// (keep waiting for the real key) or has no name (<see cref="KeyName"/>). Not validated: see
    /// <see cref="KeyCapture"/>.
    /// </summary>
    public static string? ToKeyString(int virtualKey, KeyModifiers modifiers)
    {
        if (IsModifierKey(virtualKey))
        {
            return null;
        }

        var name = KeyName(virtualKey);
        if (name is null)
        {
            return null;
        }

        var text = new StringBuilder(24);
        if ((modifiers & KeyModifiers.Ctrl) != 0)
        {
            text.Append("Ctrl+");
        }

        if ((modifiers & KeyModifiers.Alt) != 0)
        {
            text.Append("Alt+");
        }

        if ((modifiers & KeyModifiers.Shift) != 0)
        {
            text.Append("Shift+");
        }

        return text.Append(name).ToString();
    }
}
