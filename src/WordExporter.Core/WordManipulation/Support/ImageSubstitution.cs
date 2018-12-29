using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;

namespace WordExporter.Core.WordManipulation.Support
{
    public class ImageSubstitution : IDisposable
    {
        public ImageSubstitution(
            Stream imageStream)
        {
            Image = Image.FromStream(imageStream);
        }

        public Image Image { get; private set; }

        public void Dispose()
        {
            Image?.Dispose();
        }

        /// <summary>
        /// Return a memory stream containing the image saved as Jpeg.
        /// </summary>
        /// <returns></returns>
        internal Stream GetImageStream()
        {
            var memoryStream = new MemoryStream();
            ImageCodecInfo jpgEncoder = GetEncoder(ImageFormat.Jpeg);

            // Create an Encoder object based on the GUID  
            // for the Quality parameter category.  
            Encoder myEncoder = Encoder.Quality;

            // Create an EncoderParameters object.  
            // An EncoderParameters object has an array of EncoderParameter  
            // objects. In this case, there is only one  
            // EncoderParameter object in the array.  
            EncoderParameters myEncoderParameters = new EncoderParameters(1);
            myEncoderParameters.Param[0] = new EncoderParameter(myEncoder, 100L);
            Image.Save(memoryStream, jpgEncoder, myEncoderParameters);

            memoryStream.Seek(0, SeekOrigin.Begin);
            return memoryStream;
        }

        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();
            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        internal void ResizeToWidth(int width)
        {
            using (var image = Image)
            {
                Image = ResizeImageToNewWidth(image, width);
            }
        }

        private static Image ResizeImageToNewWidth(Image image, int newWidth)
        {
            Int32 newHeight = (int)(image.Height * (newWidth / (float)image.Width));
            Bitmap bmPhoto = new Bitmap(newWidth, newHeight, PixelFormat.Format24bppRgb);
            bmPhoto.SetResolution(72, 72);
            using (Graphics grPhoto = Graphics.FromImage(bmPhoto))
            {
                grPhoto.SmoothingMode = SmoothingMode.AntiAlias;
                grPhoto.InterpolationMode = InterpolationMode.HighQualityBicubic;
                grPhoto.PixelOffsetMode = PixelOffsetMode.HighQuality;
                grPhoto.CompositingQuality = CompositingQuality.HighQuality;
                grPhoto.DrawImage(image, new Rectangle(0, 0, newWidth, newHeight), 0, 0, image.Width, image.Height, GraphicsUnit.Pixel);
            }

            return bmPhoto;
        }
    }
}
