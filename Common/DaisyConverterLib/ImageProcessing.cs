
//This class contains the logic for Resampling an Image

using System;
using System.Web;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;

namespace Daisy.SaveAsDAISY.Conversion
{
    class ImageProcessing
    {

        #region Fileds
        private float? height;
        private float? width;
        private float resolution;
        private string name = string.Empty;
        private ImageFormat format;
        private float? xResolutionRate;
        private Image srcImage = null;
        #endregion

        #region Properties
        public float? Width
        {
            get { return width; }
            set { width = value; }
        }

        public float? Height
        {
            get { return height; }
            set { height = value; }
        }

        public float Resolution
        {
            get { return resolution; }
            set { resolution = value; }
        }

        public float? XResolutionRate
        {
            get { return xResolutionRate; }
            set { xResolutionRate = value; }
        }

        private float? yResolutionRate;

        public float? YResolutionRate
        {
            get { return yResolutionRate; }
            set { yResolutionRate = value; }
        }

        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        public ImageFormat Format
        {
            get { return format; }
            set { format = value; }
        }

        public Image SrcImage
        {
            get { return srcImage; }
            set { srcImage = value; }
        }
        #endregion

        #region Constructors
        public ImageProcessing()
        {

        }

        public ImageProcessing(Image SrcImage, string ImageFrmt, string ImageName)
        {
            this.Name = ImageName;
            this.SrcImage = SrcImage;
            this.Width = this.SrcImage.Width;
            this.Height = this.SrcImage.Height;
            this.Resolution = this.SrcImage.HorizontalResolution;
            this.XResolutionRate = this.Width / this.Resolution;
            this.YResolutionRate = this.Height / this.Resolution;

            ImageFormat imgFormat;
            switch (ImageFrmt.ToLower())
            {
                case "bmp":
                    imgFormat = ImageFormat.Bmp;
                    break;
                case "emf":
                    imgFormat = ImageFormat.Emf;
                    break;
                case "icon":
                    imgFormat = ImageFormat.Icon;
                    break;
                case "gif":
                    imgFormat = ImageFormat.Gif;
                    break;
                case "exif":
                    imgFormat = ImageFormat.Exif;
                    break;
                case "jpeg":
                    imgFormat = ImageFormat.Jpeg;
                    break;
                case "jpg":
                    imgFormat = ImageFormat.Jpeg;
                    break;
                case "png":
                    imgFormat = ImageFormat.Png;
                    break;
                case "tiff":
                    imgFormat = ImageFormat.Tiff;
                    break;
                case "memorybmp":
                    imgFormat = ImageFormat.MemoryBmp;
                    break;
                case "wmf":
                    imgFormat = ImageFormat.Wmf;
                    break;
                default:
                    imgFormat = ImageFormat.Jpeg;
                    break;
            }
            this.Format = imgFormat;

        }

        #endregion

        #region Methods

        public String SaveProcessedImage(string OutPutFolder)
        {
            try
            {
                Bitmap bmPhoto = new Bitmap((int)this.Width, (int)this.Height, PixelFormat.Format24bppRgb);
                bmPhoto.SetResolution(this.Resolution, this.Resolution);

                using (Graphics grPhoto = Graphics.FromImage(bmPhoto))
                {
                    grPhoto.SmoothingMode = SmoothingMode.AntiAlias;
                    grPhoto.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    grPhoto.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    grPhoto.DrawImage(this.SrcImage, new Rectangle(0, 0, (int)this.Width, (int)this.Height));
                    grPhoto.Dispose();

                }
                String tempSave = OutPutFolder + "\\" + this.Name + "." + this.Format.ToString();
                bmPhoto.Save(tempSave, this.Format);
                return UriEscape(this.Name + "." + this.Format.ToString());
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to save the image", ex);
            }
        }

        #endregion

        #region URI Images

        public String UriEscape(String imgName)
        {
            String imgNameRet;
            imgNameRet = HttpUtility.UrlEncode(imgName);
            if (imgNameRet.Contains("+"))
                imgNameRet = imgNameRet.Replace("+", "%20");
            return imgNameRet;
        }
        #endregion

    }
}
