using ImageMagick;

namespace OH_Montgomery.Services;

/// <summary>Converts base64 images to single Multi-page TIFF with CCITT Group 4 compression.</summary>
public static class ImageProcessor
{
    public static byte[] CreateMultiPageCompressedTiff(string[] base64Images)
    {
        using var collection = new MagickImageCollection();
        foreach (var b64 in base64Images)
        {
            if (string.IsNullOrWhiteSpace(b64)) continue;

            var imgBytes = Convert.FromBase64String(b64);
            var image = new MagickImage(imgBytes);
            image.ColorSpace = ColorSpace.Gray;
            image.Format = MagickFormat.Tif;
            image.Settings.Compression = CompressionMethod.Group4;
            image.Threshold(new Percentage(50));
            image.Density = new Density(300, 300);
            collection.Add(image);
        }

        using var outputStream = new MemoryStream();
        collection.Write(outputStream, MagickFormat.Tif);
        return outputStream.ToArray();
    }
}
