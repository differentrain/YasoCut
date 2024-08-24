using System.Drawing;
using System.Runtime.InteropServices;


namespace YasoCut.PInvoke
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;


        public static int Size = Marshal.SizeOf(typeof(NativeRect));

        public Rectangle ToRectangle()
        {
            return new Rectangle(Left, Top, Right - Left, Bottom - Top);
        }

        public Rectangle ToRectangle(in NativeRect realRect)
        {
            return new Rectangle(Left, Top + (Bottom - Top - realRect.Bottom), realRect.Right, realRect.Bottom);
        }
    }
}
