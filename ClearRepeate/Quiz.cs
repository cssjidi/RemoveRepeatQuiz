using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace RemoveRepeat
{
    public class Quiz
    {
        /// <summary>
        /// Title
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// Image 
        /// </summary>
        public List<string> Image { get; set; }
        /// <summary>
        /// Picture
        /// </summary>
        public string[] Picture { get; set; }

        #region 判断图片是否一致
        /// <summary>
        /// 判断图片是否一致
        /// </summary>
        /// <param name="img">图片一
        /// <param name="bmp">图片二
        /// <returns>是否一致</returns>
        public bool IsSamePicture(Bitmap img, Bitmap bmp)
        {
            //大小一致
            if (img.Width == bmp.Width && img.Height == bmp.Height)
            {
                //将图片一锁定到内存
                BitmapData imgData_i = img.LockBits(new Rectangle(0, 0, img.Width, img.Height), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                IntPtr ipr_i = imgData_i.Scan0;
                int length_i = imgData_i.Width * imgData_i.Height * 3;
                byte[] imgValue_i = new byte[length_i];
                Marshal.Copy(ipr_i, imgValue_i, 0, length_i);
                img.UnlockBits(imgData_i);
                //将图片二锁定到内存
                BitmapData imgData_b = img.LockBits(new Rectangle(0, 0, img.Width, img.Height), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
                IntPtr ipr_b = imgData_b.Scan0;
                int length_b = imgData_b.Width * imgData_b.Height * 3;
                byte[] imgValue_b = new byte[length_b];
                Marshal.Copy(ipr_b, imgValue_b, 0, length_b);
                img.UnlockBits(imgData_b);
                //长度不相同
                if (length_i != length_b)
                {
                    return false;
                }
                else
                {
                    //循环判断值
                    for (int i = 0; i < length_i; i++)
                    {
                        //不一致
                        if (imgValue_i[i] != imgValue_b[i])
                        {
                            return false;
                        }
                    }
                    return true;
                }
            }
            else
            {
                return false;
            }
        }
        #endregion
    }
}
