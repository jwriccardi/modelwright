using System;
using System.Runtime.InteropServices;
using System.Text;
using Modelwright.Core.Keys;

namespace Modelwright.AddIn;

/// <summary>
/// What a captured key press means on the active keyboard layout, for the settings dialog's Capture button.
/// <see cref="KeyCapture"/> names punctuation keys by their US-layout character; on other layouts these helpers
/// use the character the layout puts on the key, and warn when Ctrl+Alt (AltGr) with the key types a character.
/// </summary>
internal static class KeyboardLayout
{
    private const uint MapVkToScanCode = 0; // MAPVK_VK_TO_VSC
    private const uint MapVkToChar = 2; // MAPVK_VK_TO_CHAR
    private const uint KeepKeyboardState = 0x4; // ToUnicodeEx: do not change the keyboard state (Windows 10 1607+)
    private const int VkShift = 0x10;
    private const int VkControl = 0x11;
    private const int VkMenu = 0x12;
    private const int VkLShift = 0xA0;
    private const int VkLControl = 0xA2;
    private const int VkRMenu = 0xA5;

    /// <summary>
    /// The key string for <paramref name="virtualKey"/> with <paramref name="modifiers"/> on the active layout, given
    /// <paramref name="usKeyString"/>, what <see cref="KeyCapture.ToKeyString"/> made of it. For a punctuation
    /// (<c>VK_OEM_*</c>) key whose character on this layout differs from its US one, the layout's character is used
    /// if a shortcut can use it; if not, returns null with the reason in <paramref name="problem"/>. Any other key is
    /// returned unchanged (null <paramref name="problem"/>).
    /// </summary>
    public static string? ToLayoutKeyString(int virtualKey, KeyModifiers modifiers, string? usKeyString, out string? problem)
    {
        problem = null;
        if (!IsOemKey(virtualKey))
        {
            return usKeyString;
        }

        // The low word is the unshifted character; the top bit marks a dead key (an accent), which still names the key.
        var layoutChar = (char)(MapVirtualKeyEx((uint)virtualKey, MapVkToChar, GetKeyboardLayout(0)) & 0xFFFF);
        var usName = KeyCapture.KeyName(virtualKey);
        if (layoutChar == '\0' || (usName is not null && usName.Length == 1 && usName[0] == layoutChar))
        {
            return usKeyString;
        }

        try
        {
            KeyChord.Parse("Ctrl+" + layoutChar); // Is the character itself a key a shortcut can use?
        }
        catch (FormatException)
        {
            problem =
                $"That key types '{layoutChar}' on your keyboard layout, which a shortcut can't use. Use a letter, " +
                "digit or punctuation key found on a US keyboard, F1-F12, an arrow, PgUp, PgDn, Home, End, Ins or Del.";
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

        return text.Append(char.ToUpperInvariant(layoutChar)).ToString();
    }

    /// <summary>
    /// A warning if <paramref name="modifiers"/> include Ctrl and Alt and, on the active layout, that combination
    /// with <paramref name="virtualKey"/> types a character (AltGr is Ctrl+Alt to Windows): binding it would stop
    /// people typing that character. Null otherwise.
    /// </summary>
    public static string? AltGrWarning(int virtualKey, KeyModifiers modifiers)
    {
        if ((modifiers & KeyModifiers.Ctrl) == 0 || (modifiers & KeyModifiers.Alt) == 0)
        {
            return null;
        }

        var state = new byte[256];
        state[VkControl] = state[VkLControl] = 0x80;
        state[VkMenu] = state[VkRMenu] = 0x80;
        if ((modifiers & KeyModifiers.Shift) != 0)
        {
            state[VkShift] = state[VkLShift] = 0x80;
        }

        var layout = GetKeyboardLayout(0);
        var scanCode = MapVirtualKeyEx((uint)virtualKey, MapVkToScanCode, layout);
        var buffer = new StringBuilder(8);
        var count = ToUnicodeEx((uint)virtualKey, scanCode, state, buffer, buffer.Capacity, KeepKeyboardState, layout);
        if (count < 0)
        {
            // A dead key: before Windows 10 1607 the flag is ignored and the accent is now pending; this clears it.
            ToUnicodeEx((uint)virtualKey, scanCode, state, new StringBuilder(8), 8, KeepKeyboardState, layout);
        }

        if (count == 0 || (count > 0 && char.IsControl(buffer[0])))
        {
            return null;
        }

        var typed = count > 0 ? $" (this one types '{buffer.ToString(0, count)}')" : string.Empty;
        return "AltGr/Ctrl+Alt combinations type characters on this keyboard layout; binding it will block typing " +
            "that character" + typed + ".";
    }

    /// <summary>VK_OEM_1 to VK_OEM_3, VK_OEM_4 to VK_OEM_8 and VK_OEM_102: the keys layouts assign punctuation to.</summary>
    private static bool IsOemKey(int virtualKey) =>
        (virtualKey >= 0xBA && virtualKey <= 0xC0) || (virtualKey >= 0xDB && virtualKey <= 0xDF) || virtualKey == 0xE2;

    [DllImport("user32.dll")]
    private static extern IntPtr GetKeyboardLayout(uint threadId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern uint MapVirtualKeyEx(uint code, uint mapType, IntPtr layout);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int ToUnicodeEx(
        uint virtualKey,
        uint scanCode,
        byte[] keyState,
        [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder buffer,
        int bufferSize,
        uint flags,
        IntPtr layout);
}
