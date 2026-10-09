using System;
using System.Globalization;

namespace Modelwright.Core.Trace;

/// <summary>A window's or monitor work area's bounds in screen pixels (Win32 coordinates: y grows downwards).</summary>
public readonly struct WindowRect : IEquatable<WindowRect>
{
    /// <summary>Creates a rectangle; a negative size is taken as zero.</summary>
    public WindowRect(int left, int top, int width, int height)
    {
        Left = left;
        Top = top;
        Width = Math.Max(0, width);
        Height = Math.Max(0, height);
    }

    /// <summary>The left edge.</summary>
    public int Left { get; }

    /// <summary>The top edge.</summary>
    public int Top { get; }

    /// <summary>The width.</summary>
    public int Width { get; }

    /// <summary>The height.</summary>
    public int Height { get; }

    /// <summary>One past the right edge (<see cref="Left"/> + <see cref="Width"/>).</summary>
    public int Right => Left + Width;

    /// <summary>One past the bottom edge (<see cref="Top"/> + <see cref="Height"/>).</summary>
    public int Bottom => Top + Height;

    /// <summary>The same size at another position.</summary>
    public WindowRect MoveTo(int left, int top) => new WindowRect(left, top, Width, Height);

    /// <summary>The overlap's area (0 if none), in square pixels.</summary>
    public long OverlapArea(WindowRect other)
    {
        long width = Math.Min(Right, other.Right) - Math.Max(Left, other.Left);
        long height = Math.Min(Bottom, other.Bottom) - Math.Max(Top, other.Top);
        return width <= 0 || height <= 0 ? 0 : width * height;
    }

    /// <inheritdoc />
    public bool Equals(WindowRect other) =>
        Left == other.Left && Top == other.Top && Width == other.Width && Height == other.Height;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is WindowRect other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        unchecked
        {
            return (((((Left * 397) ^ Top) * 397) ^ Width) * 397) ^ Height;
        }
    }

    /// <summary>Value equality.</summary>
    public static bool operator ==(WindowRect left, WindowRect right) => left.Equals(right);

    /// <summary>Value inequality.</summary>
    public static bool operator !=(WindowRect left, WindowRect right) => !left.Equals(right);

    /// <summary><c>(left,top) widthxheight</c>, for the log.</summary>
    public override string ToString() =>
        string.Format(CultureInfo.InvariantCulture, "({0},{1}) {2}x{3}", Left, Top, Width, Height);
}
