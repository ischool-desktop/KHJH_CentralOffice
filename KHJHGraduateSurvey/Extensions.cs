using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JH_KH_GraduateSurvey
{
    /// <summary>
    /// 常用的函式功能
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// 將日期時間轉為當日的開始
        /// </summary>
        /// <param name="DateTime"></param>
        /// <returns></returns>
        public static DateTime ToDayStart(this DateTime DateTime)
        {
            return new DateTime(DateTime.Year, DateTime.Month, DateTime.Day, 00, 00, 01);
        }

        /// <summary>
        /// 將日期時間轉為當日的結束
        /// </summary>
        /// <param name="DateTime"></param>
        /// <returns></returns>
        public static DateTime ToDayEnd(this DateTime DateTime)
        {
            return new DateTime(DateTime.Year, DateTime.Month, DateTime.Day, 23, 59, 59);
        }
    }
}