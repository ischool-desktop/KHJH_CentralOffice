using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DesktopLib;
using FISCA.UDT;
using System.Linq;

namespace KHJHCentralOffice
{
    public partial class GraduateSurveyVagrant : DetailContentImproved
    {
        private List<VagrantStatistics> VargantSats { get; set; }

        public GraduateSurveyVagrant()
        {
            InitializeComponent();
            Group = "未升學未就業統計";
        }

        protected override void OnInitializeComplete(Exception error)
        {

        }

        protected override void OnSaveData()
        {

        }

        protected override void OnPrimaryKeyChangedAsync()
        {          
            VargantSats = Utility.AccessHelper
                .Select<VagrantStatistics>(string.Format("ref_school_id='{0}'", PrimaryKey))
                .OrderBy(x=>x.SurveyYear)
                .ToList();
        }

        private void ResolveUrl()
        {
            //PhysicalUrl = string.Empty;
            //if (SchoolData != null)
            //{
            //    AccessPoint ap;
            //    if (AccessPoint.TryParse(SchoolData.DSNS, out ap))
            //        PhysicalUrl = ap.Url;
            //}
        }

        protected override void OnPrimaryKeyChangedComplete(Exception error)
        {
            if (VargantSats != null)
            {
                BeginChangeControlData();

                grdVagrant.Rows.Clear();

                foreach(VagrantStatistics Sat in VargantSats)
                {
                    grdVagrant.Rows.Add(
                        Sat.SurveyYear, 
                        Sat.InJob, 
                        Sat.InSchool, 
                        Sat.PrepareSchool, 
                        Sat.PrepareJob, 
                        Sat.InTraining, 
                        Sat.InHome, 
                        Sat.NoPlan, 
                        Sat.DisAppearance, 
                        Sat.Other);
                }

                Task task = Task.Factory.StartNew(() =>
                {
                    ResolveUrl();
                }, CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default);

                task.ContinueWith(x =>
                {
                    //txtPhysicalUrl.Text = PhysicalUrl;
                }, TaskScheduler.FromCurrentSynchronizationContext());

                ResetDirtyStatus();
            }
            else
                throw new Exception("無查資料：" + PrimaryKey);
        }

        private void BasicInfoItem_Load(object sender, EventArgs e)
        {
            InitDetailContent();
        }
    }
}
