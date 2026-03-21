namespace OneInk
{
    using System;
    using System.Drawing;
    using System.Drawing.Imaging;
    using System.IO;
    using System.Runtime.InteropServices.ComTypes;

    internal static class BitmapExtensions
    {
        public static IStream GetReadOnlyStream(this Bitmap bitmap)
        {
            try
            {
                var memory = new MemoryStream();
                bitmap.Save(memory, ImageFormat.Png);
                memory.Position = 0;
                return new ReadOnlyStream(memory);
            }
            catch
            {
                return null;
            }
        }
    }
}
