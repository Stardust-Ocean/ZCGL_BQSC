using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using ZXing;
using ZXing.Common;
using ZXing.QrCode;


namespace PsjgWeb.code
{
    public static class MakeZxing
    {
        /// <summary>
        /// 生产二维码的图片的字节流，值，logo文件，宽度，高度，默认180X180
        /// </summary>
        /// <param name="codeResult">值</param>
        /// <param name="logoFile"></param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <returns></returns>
        public static byte[] GetMakeQrCodeByte(string codeResult, string logoFile = null, int width = 180, int height = 180)
        {
            return MakeQrCodeMemoryStream(codeResult, logoFile, width, height).ToArray();
        }

        /// <summary>
        /// 生产二维码的图片的流，值，logo文件，宽度，高度，默认180X180
        /// </summary>
        /// <param name="codeResult">值</param>
        /// <param name="logoFile"></param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <returns></returns>
        public static MemoryStream GetMakeQrCodeMemoryStream(string codeResult, string logoFile = null, int width = 180, int height = 180)
        {
            return MakeQrCodeMemoryStream(codeResult, logoFile, width, height);
        }

        /// <summary>
        /// 将指定logo文件加入到BitMap中
        /// </summary>
        /// <param name="bitmap"></param>
        /// <param name="logoFile"></param>
        /// <returns></returns>
        private static Bitmap AddLogoToBitmap(Image bitmap, string logoFile)
        {
            Bitmap retBitmap = new Bitmap(bitmap);
            if (string.IsNullOrWhiteSpace(logoFile)) { return retBitmap; }
            if (!File.Exists(logoFile)) { return retBitmap; }
            Image logo = Image.FromFile(logoFile);
            Graphics g = Graphics.FromImage(retBitmap);

            //draw center logo
            g.DrawImage(logo, (bitmap.Width - logo.Width) / 2, (bitmap.Width - logo.Width) / 2, logo.Width, logo.Width);
            g.Dispose();
            return retBitmap;
        }

        /// <summary>
        /// 生产二维码的图片的字节流，值，logo文件绝对路径，宽度，高度
        /// </summary>
        /// <param name="codeResult">值</param>
        /// <param name="logoFile"></param>
        /// <param name="width">宽度</param>
        /// <param name="height">高度</param>
        /// <returns></returns>
        private static MemoryStream MakeQrCodeMemoryStream(string codeResult, string logoFile = null, int width = 180, int height = 180)
        {
            width = width == 0 ? 180 : width;
            height = height == 0 ? 180 : height;
            EncodingOptions options = new QrCodeEncodingOptions
            {
                DisableECI = true,
                CharacterSet = "UTF-8",
                Width = width,
                Height = height
            };

            BarcodeWriter writer = new BarcodeWriter { Format = BarcodeFormat.QR_CODE, Options = options };
            Bitmap bitmap = writer.Write(codeResult);
            bitmap = AddLogoToBitmap(bitmap, logoFile);
            MemoryStream memoryStream = new MemoryStream();
            bitmap.Save(memoryStream, ImageFormat.Gif);
            return memoryStream;
        }

    }
}
