using System.Drawing.Imaging;

namespace YasoCut.Internals
{
    internal static class ImageFormatTypeExtensions
    {
        public static ImageFormat GetImageFormat(this ImageFormatType type)
        {
            switch (type)
            {
                case ImageFormatType.Bmp:
                    return ImageFormat.Bmp;
                case ImageFormatType.Gif:
                    return ImageFormat.Gif;
                case ImageFormatType.Jpeg:
                    return ImageFormat.Jpeg;
                case ImageFormatType.Png:
                    return ImageFormat.Png;
                case ImageFormatType.Icon:
                    return ImageFormat.Icon;
                case ImageFormatType.Tiff:
                    return ImageFormat.Tiff;
                case ImageFormatType.Exif:
                    return ImageFormat.Exif;
                case ImageFormatType.Emf:
                    return ImageFormat.Emf;
                case ImageFormatType.Wmf:
                    return ImageFormat.Wmf;
                default:
                    return null;
            }
        }

        public static string GetImageExtensionName(this ImageFormatType type)
        {
            switch (type)
            {
                case ImageFormatType.Bmp:
                    return "bmp";
                case ImageFormatType.Gif:
                    return "gif";
                case ImageFormatType.Jpeg:
                    return "jpeg";
                case ImageFormatType.Png:
                    return "png";
                case ImageFormatType.Icon:
                    return "ico";
                case ImageFormatType.Tiff:
                    return "tiff";
                case ImageFormatType.Exif:
                    return "jpeg";
                case ImageFormatType.Emf:
                    return "emf";
                case ImageFormatType.Wmf:
                    return "wmf";
                default:
                    return null;
            }
        }
    }
}
