using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using DesktopLib;
using FISCA;
using FISCA.Authentication;
using FISCA.Permission;
using FISCA.Presentation;
using FISCA.UDT;

namespace KHJHCentralOffice
{
    public static class Program
    {
        public static UserConfigManager User { get; private set; }

        public static ConfigurationManager App { get; private set; }

        public static ConfigurationManager Global { get; private set; }

        public static NLDPanel MainPanel { get; private set; }

        public static DynamicCache GlobalSchoolCache { get; private set; }

        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [MainMethod]
        public static void Main()
        {
            #region 模組啟用先同步Schmea
            //K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["調代課UDT載入設定"];

            //bool checkClubUDT = false;
            //string name = "調代課UDT_20131008";

            ////如果尚無設定值,預設為
            //if (string.IsNullOrEmpty(cd[name]))
            //{
            //    cd[name] = "false";
            //}
            ////檢查是否為布林
            //bool.TryParse(cd[name], out checkClubUDT);

            //if (!checkClubUDT)
            //{

			//ServerModule.AutoManaged("http://module.ischool.com.tw/module/89/KHCentralOffice/udm.xml");

            SchemaManager Manager = new SchemaManager(FISCA.Authentication.DSAServices.DefaultConnection);

            Manager.SyncSchema(new School());
            Manager.SyncSchema(new ApproachStatistics());
            Manager.SyncSchema(new VagrantStatistics());
            Manager.SyncSchema(new SchoolLog());

                //cd[name] = "true";
                //cd.Save();
            //}
            #endregion

            FISCA.Presentation.MotherForm.StartMenu["安全性"]["權限管理"].Click += (sender, e) => new FISCA.Permission.UI.RoleManager().ShowDialog();

            DSAServices.AutoDisplayLoadingMessageOnMotherForm();

            GlobalSchoolCache = new DynamicCache(); //建立一個空的快取。

            InitAsposeLicense();
            InitStartMenu();
            InitConfigurationStorage();
            InitMainPanel();

            //MainPanel.ListPaneContexMenu["執行 SQL 並匯出"].Click += delegate
            //{
            //    new ExportQueryData().Export();
            //};

            new FieldManager();
            new DetailItems();
            //new RibbonButtons();
            //new ImportExport();//匯入學校資料

            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["開放時間"].Image = Properties.Resources.school_events_config_128;
            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["開放時間"].Size = RibbonBarButton.MenuButtonSize.Medium;
            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["開放時間"].Click += (sender, e) => new OpenTime().ShowDialog();
            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["開放時間"].Enable = Permissions.開放時間權限;

            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["報表"].Image = Properties.Resources.paste_64;
			Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["報表"].Size = RibbonBarButton.MenuButtonSize.Medium;

			#region 匯出畢業學生進路統計分析資料
			Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["匯出"].Image = Properties.Resources.Export_Image;
			Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["匯出"].Size = RibbonBarButton.MenuButtonSize.Medium;
			Catalog button_Approach_Export = RoleAclSource.Instance["學校"]["功能按鈕"];
			button_Approach_Export.Add(new RibbonFeature("button_Approach_Export", "匯出畢業學生進路統計分析資料"));
			Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["匯出"]["畢業學生進路統計分析資料"].Enable = UserAcl.Current["button_Approach_Export"].Executable;
			Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["匯出"]["畢業學生進路統計分析資料"].Click += delegate
			{
				(new Approach_Export()).ShowDialog();
			};
			#endregion

            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["報表"]["畢業學生進路統計表"].Click += (sender, e) => new Approach_Report("國中畢業學生進路調查填報表格", Accessor.ApproachReportTemplate.ReportType.統計表).ShowDialog();
            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["報表"]["畢業學生進路統計表"].Enable = Permissions.畢業學生進路統計表權限;

            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["報表"]["畢業學生進路複核表"].Click += (sender, e) => new Approach_Report("國中畢業學生進路調查填報複核表", Accessor.ApproachReportTemplate.ReportType.複核表).ShowDialog();
            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["報表"]["畢業學生進路複核表"].Enable = Permissions.畢業學生進路複核表權限;

            //Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["報表"]["畢業未升學未就業學生動向"].Click += (sender, e) => new UnApproach_Report("國中畢業未升學未就業學生動向", Properties.Resources.國中畢業未升學未就業學生動向).ShowDialog();
            //Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["報表"]["畢業未升學未就業學生動向"].Enable = Permissions.畢業未升學未就業學生動向權限;

            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["未上傳學校"].Size = RibbonBarButton.MenuButtonSize.Medium;
            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["未上傳學校"].Image = Properties.Resources.school_search_128;
            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["未上傳學校"].Click += (sender, e) => new UnApproach_Check().ShowDialog();
            Program.MainPanel.RibbonBarItems["畢業學生進路調查"]["未上傳學校"].Enable = Permissions.未上傳學校權限;

            FISCA.Permission.Catalog AdminCatalog = FISCA.Permission.RoleAclSource.Instance["畢業學生進路調查"]["功能按鈕"];
            AdminCatalog.Add(new RibbonFeature(Permissions.未上傳學校, "未上傳學校"));
            AdminCatalog.Add(new RibbonFeature(Permissions.畢業未升學未就業學生動向, "畢業未升學未就業學生動向"));
            AdminCatalog.Add(new RibbonFeature(Permissions.畢業學生進路統計表, "畢業學生進路統計表"));
            AdminCatalog.Add(new RibbonFeature(Permissions.畢業學生進路複核表, "畢業學生進路複核表"));
            AdminCatalog.Add(new RibbonFeature(Permissions.開放時間, "開放時間"));

            FISCA.Permission.Catalog DetailCatalog = FISCA.Permission.RoleAclSource.Instance["畢業學生進路調查"]["資料項目"];
            DetailCatalog.Add(new DetailItemFeature(Permissions.學校基本資料, "學校基本資料"));
            DetailCatalog.Add(new DetailItemFeature(Permissions.學校進路統計, "學校進路統計"));

            RefreshFilteredSource();

            FISCA.Presentation.MotherForm.Form.Text = GetTitleText();
        }

        private static void InitMainPanel()
        {
            MainPanel = new NLDPanel();
            MainPanel.Group = "學校";
            MainPanel.SetDescriptionPaneBulider<DetailItemDescription>();

            InitBasicSearch();

            MotherForm.AddPanel(MainPanel);
            MainPanel.AddView(new DefaultView());
        }

        private static void InitBasicSearch()
        {
            MainPanel.Search += delegate(object sender, SearchEventArgs args)
            {
                string cond = args.Condition;
                foreach (string each in GlobalSchoolCache.PrimaryKeys)
                {
                    string text = GlobalSchoolCache[each].Title;
                    if (text.IndexOf(cond) >= 0)
                    {
                        args.Result.Add(each);
                        continue;
                    }

                    text = GlobalSchoolCache[each].DSNS;
                    if (text.IndexOf(cond) >= 0)
                    {
                        args.Result.Add(each);
                        continue;
                    }

                    text = GlobalSchoolCache[each].Group;
                    if (text.IndexOf(cond) >= 0)
                    {
                        args.Result.Add(each);
                        continue;
                    }

                    text = GlobalSchoolCache[each].Comment;
                    if (text.IndexOf(cond) >= 0)
                    {
                        args.Result.Add(each);
                        continue;
                    }
                }
            };
        }

        internal static void RefreshFilteredSource()
        {
            RefreshFilteredSource(null);
        }

        internal static void RefreshFilteredSource(Action callback)
        {
            List<string> schoolids = new List<string>();
            Task task = Task.Factory.StartNew(() =>
            {
                AccessHelper access = new AccessHelper();

                List<School> schools = access.Select<School>();

                foreach (School school in schools)
                {
                    schoolids.Add(school.UID);
                    GlobalSchoolCache.FillProperty(school.UID, "Title", school.Title);
                    GlobalSchoolCache.FillProperty(school.UID, "DSNS", school.DSNS);
                    GlobalSchoolCache.FillProperty(school.UID, "Group", school.Group);
                    GlobalSchoolCache.FillProperty(school.UID, "Comment", school.Comment);
                }
            }, CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default);

            task.ContinueWith((x) =>
            {
                MainPanel.SetFilteredSource(schoolids);

                if (callback != null)
                    callback();
            }, TaskScheduler.Default);
        }

        private static void InitConfigurationStorage()
        {
            User = new UserConfigManager(new ConfigProvider_User(), DSAServices.UserAccount);
            App = new ConfigurationManager(new ConfigProvider_App());
            Global = new ConfigurationManager(new ConfigProvider_Global());
        }

        private static void InitStartMenu()
        {
            MotherForm.StartMenu["安全性"]["使用者管理"].Click += delegate
            {
                new FISCA.Permission.UI.UserManager().ShowDialog();
            };

            FISCA.Presentation.MotherForm.StartMenu["重新登入"].BeginGroup = true;
            FISCA.Presentation.MotherForm.StartMenu["重新登入"].Click += delegate
            {
                Application.Restart();
            };
        }

        private static void InitAsposeLicense()
        {
            //設定 ASPOSE 元件的 License。
            System.IO.Stream stream = new System.IO.MemoryStream(Properties.Resources.Aspose_Total);
            new Aspose.Cells.License().SetLicense(stream);

            stream.Seek(0, SeekOrigin.Begin);
            new Aspose.Words.License().SetLicense(stream);
        }

        /// <summary>
        /// 設定ischool標頭
        /// </summary>
        private static string GetTitleText()
        {
            string user = DSAServices.UserAccount;

            string version = "0.0.0.0";
            try
            {
                string path = Path.Combine(Application.StartupPath, "release.xml");
                XmlDocument doc = new XmlDocument();
                doc.Load(path);
                version = doc.DocumentElement.GetAttribute("Version");
            }
            catch (Exception) { }

            return string.Format("ischool 中央管理系統〈FISCA：{0}〉〈{1}〉", version, user);
        }

        internal static string TrimModuleServerUrl(string urlPart)
        {
            string part = urlPart.Replace("https://module.ischool.com.tw/module", "~");
            part = part.Replace("http://module.ischool.com.tw/module", "~");
            part = part.Replace("http://module.ischool.com.tw:8080/module", "~");
            part = part.Replace("https://module.ischool.com.tw:8080/module", "~");
            return part;
        }
    }
}
