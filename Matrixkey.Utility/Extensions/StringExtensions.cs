using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Matrixkey.Utility.Extensions
{
    public static class StringExtensions
    {
        /// <summary>
        /// 将字符串转换为其所表示的 int 值。
        /// </summary>
        /// <param name="value">要转换的字符串</param>
        /// <returns></returns>
        public static int ToInt(this string value)
        {
            return ToInt(value, 0);
        }

        /// <summary>
        /// 将字符串转换为其所表示的 int 值。参数表示转换不成功时返回的默认值。
        /// </summary>
        /// <param name="value">要转换的字符串</param>
        /// <param name="defaultValue">转换失败时要返回的默认值</param>
        /// <returns></returns>
        public static int ToInt(this string value, int defaultValue)
        {
            int ret;
            return int.TryParse(value, out ret) ? ret : defaultValue;
        }

        /// <summary>
        /// 将字符串转换为其所表示的 decimal 值。
        /// </summary>
        /// <param name="value">要转换的字符串</param>
        /// <returns></returns>
        public static decimal ToDecimal(this string value)
        {
            return ToDecimal(value, 0M);
        }

        /// <summary>
        /// 将字符串转换为其所表示的 decimal 值。
        /// </summary>
        /// <param name="value">要转换的字符串</param>
        /// <param name="defaultValue">转换失败时要返回的默认值</param>
        /// <returns></returns>
        public static decimal ToDecimal(this string value, decimal defaultValue)
        {
            decimal ret;
            return decimal.TryParse(value, out ret) ? ret : defaultValue;
        }

        /// <summary>
        /// 获取字符串的实际字节数
        /// </summary>
        /// <param name="value">要获取长度的字符串</param>
        /// <returns></returns>
        public static long GetByteLength(this string value)
        {
            if (value.Equals(string.Empty))
            { return 0; }
            int strlen = 0;
            ASCIIEncoding strData = new ASCIIEncoding();
            //将字符串转换为ASCII编码的字节数字
            byte[] strBytes = strData.GetBytes(value);
            for (int i = 0; i <= strBytes.Length - 1; i++)
            {
                if (strBytes[i] == 63)  //中文都将编码为ASCII编码63,即"?"号
                { strlen++; }
                strlen++;
            }
            return strlen;
        }
    }
}
