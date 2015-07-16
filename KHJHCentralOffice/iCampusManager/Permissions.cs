
namespace KHJHCentralOffice
{
    class Permissions
    {
        public const string 開放時間 = "ed43c373-d087-4a14-aebb-0f1fc9b5129f";

        public static bool 開放時間權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[開放時間].Executable;
            }
        }

        public const string 畢業學生進路統計表 = "01cd6092-70f6-498c-be31-5ca1b2756642";

        public static bool 畢業學生進路統計表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[畢業學生進路統計表].Executable;
            }
        }

        public const string 畢業學生進路複核表 = "90010945-7f33-441b-92b9-2359f001bcaa";

        public static bool 畢業學生進路複核表權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[畢業學生進路複核表].Executable;
            }
        }

        public const string 畢業未升學未就業學生動向 = "61ed6f78-f27a-478b-8f4e-0bd0d2536216";

        public static bool 畢業未升學未就業學生動向權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[畢業未升學未就業學生動向].Executable;
            }
        }

        public const string 未上傳學校 = "8ec76011-6d72-4b2c-bed8-43812ce00caa";

        public static bool 未上傳學校權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[未上傳學校].Executable;
            }
        }


        public const string 學校基本資料 = "f365f269-8e73-49ca-8354-92f4eb2c0084";

        public static bool 學校基本資料權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學校基本資料].Executable;
            }
        }


        public const string 學校進路統計 = "d0077a49-80a6-45d6-9c6a-b54c470f64cb";

        public static bool 學校進路統計權限
        {
            get
            {
                return FISCA.Permission.UserAcl.Current[學校進路統計].Executable;
            }
        }
    }
}