using FISCA.UDT;

namespace KHJHCentralOffice
{
    /// <summary>
    /// 常用函式庫
    /// </summary>
    public class Utility
    {
        private static AccessHelper mHelper = null;

        /// <summary>
        /// 取得AccessHelper
        /// </summary>
        /// <returns></returns>
        public static AccessHelper AccessHelper
        {
            get
            {
                if (mHelper == null)
                    mHelper = new AccessHelper();

                return mHelper;
            }
        }
    }
}