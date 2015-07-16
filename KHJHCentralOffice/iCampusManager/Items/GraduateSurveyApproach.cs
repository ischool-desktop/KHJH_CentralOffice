using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using DesktopLib;
using DevComponents.DotNetBar.Controls;
using FISCA.UDT;
using FCode = FISCA.Permission.FeatureCodeAttribute;

namespace KHJHCentralOffice
{
    [FCode(Permissions.學校進路統計, "學校進路統計")]
    public partial class GraduateSurveyApproach : DetailContentImproved
    {
        private List<ApproachStatistics> ApproachSats { get; set; }

        private string PhysicalUrl { get; set; }

        public GraduateSurveyApproach()
        {
            InitializeComponent();
            Group = "學校進路統計";
        }

        protected override void OnInitializeComplete(Exception error)
        {

        }

        protected override void OnSaveData()
        {

        }

        protected override void OnPrimaryKeyChangedAsync()
        {            AccessHelper access = new AccessHelper();
            ApproachSats = Utility.AccessHelper
                .Select<ApproachStatistics>(string.Format("ref_school_id={0}", PrimaryKey));
        }

        private void SetControl(int SchoolYear)
        {
            foreach (Control vControl in this.Controls)
            {
                if (vControl is TextBoxX)
                {
                    TextBoxX vTextBox = vControl as TextBoxX;
                    vTextBox.Text = string.Empty;
                }
            }

            for (int i = 0; i < cmbSurveyYear.Items.Count; i++)
            {
                if (cmbSurveyYear.Items[i].Equals(SchoolYear))
                {
                    cmbSurveyYear.SelectedIndex = i;
                    break;
                }
            }

            ApproachStatistics ApproachSat = ApproachSats
                .Find(x => x.SurveyYear.Equals(SchoolYear));

            if (ApproachSat != null)
            {

            }
        }

        protected override void OnPrimaryKeyChangedComplete(Exception error)
        {
            grdContent.Rows.Clear();
            cmbSurveyYear.Items.Clear();
            cmbSurveyYear.Text = string.Empty;

            if (ApproachSats != null)
            {
                BeginChangeControlData();

                List<int> SurveyYears = ApproachSats
                    .Select(x => x.SurveyYear)
                    .ToList();

                SurveyYears.ForEach(x => cmbSurveyYear.Items.Add(x));
                cmbSurveyYear.SelectedIndex = cmbSurveyYear.Items.Count - 1;

                ResetDirtyStatus();
            }
            else
                throw new Exception("無查資料：" + PrimaryKey);
        }

        private void BasicInfoItem_Load(object sender, EventArgs e)
        {
            InitDetailContent();
        }

        private void cmbSurveyYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            string Value = "" + cmbSurveyYear.SelectedItem;

            grdContent.Rows.Clear();

            if (!string.IsNullOrEmpty(Value))
            {
                Dictionary<string, string> LookupValues = new Dictionary<string, string>();

                LookupValues.Add("A1", "全校畢業學生升學與就業情形 - 畢業學生總數");
                LookupValues.Add("B1", "全校畢業學生升學與就業情形 - 升學學生數");
                LookupValues.Add("B1-A1", "全校畢業學生升學與就業情形 - 升學比率");
                LookupValues.Add("C1", "全校畢業學生升學與就業情形 - 就業學生數");
                LookupValues.Add("C1-A1", "全校畢業學生升學與就業情形 - 就業比率");
                LookupValues.Add("D1", "全校畢業學生升學與就業情形 - 未升學未就業學生數");
                LookupValues.Add("D1-A1", "全校畢業學生升學與就業情形 - 未升學未就業比率");

                LookupValues.Add("A2", "全校畢業學生升學之就讀學校情形 - 升學學生總數");
                LookupValues.Add("B2", "全校畢業學生升學之就讀學校情形 - 公立高中學生數");
                LookupValues.Add("B2-A2", "全校畢業學生升學之就讀學校情形 - 公立高中比率");
                LookupValues.Add("C2", "全校畢業學生升學之就讀學校情形 - 私立高中學生數");
                LookupValues.Add("C2-A2", "全校畢業學生升學之就讀學校情形 - 私立高中比率");
                LookupValues.Add("D2", "全校畢業學生升學之就讀學校情形 - 公立高職學生數");
                LookupValues.Add("D2-A2", "全校畢業學生升學之就讀學校情形 - 公立高職比率");
                LookupValues.Add("E2", "全校畢業學生升學之就讀學校情形 - 私立高職學生數");
                LookupValues.Add("E2-A2", "全校畢業學生升學之就讀學校情形 - 私立高職比率");
                LookupValues.Add("F2", "全校畢業學生升學之就讀學校情形 - 五專學生數");
                LookupValues.Add("F2-A2", "全校畢業學生升學之就讀學校情形 - 五專比率");
                LookupValues.Add("G2", "全校畢業學生升學之就讀學校情形 - 軍事學校學生數");
                LookupValues.Add("G2-A2", "全校畢業學生升學之就讀學校情形 - 軍事學校比率");
                LookupValues.Add("H2", "全校畢業學生升學之就讀學校情形 - 赴國外或大陸就學學生數");
                LookupValues.Add("H2-A2", "全校畢業學生升學之就讀學校情形 - 赴國外或大陸就學比率");
                LookupValues.Add("I2", "全校畢業學生升學之就讀學校情形 - 其他學生數");
                LookupValues.Add("I2-A2", "全校畢業學生升學之就讀學校情形 - 其他比率");

                LookupValues.Add("A3", "全校畢業學生升學就讀學校之學制別 - 升學學生總數");
                LookupValues.Add("B3", "全校畢業學生升學就讀學校之學制別 - 職業類科學生數");
                LookupValues.Add("B3-A3", "全校畢業學生升學就讀學校之學制別 - 職業類科比率");
                LookupValues.Add("C3", "全校畢業學生升學就讀學校之學制別 - 綜合高中學生數");
                LookupValues.Add("C3-A3", "全校畢業學生升學就讀學校之學制別 - 綜合高中比率");
                LookupValues.Add("D3", "全校畢業學生升學就讀學校之學制別 - 普通高中學生數");
                LookupValues.Add("D3-A3", "全校畢業學生升學就讀學校之學制別 - 普通高中比率");
                LookupValues.Add("E3", "全校畢業學生升學就讀學校之學制別 - 建教合作班學生數");
                LookupValues.Add("E3-A3", "全校畢業學生升學就讀學校之學制別 - 建教合作班比率");
                LookupValues.Add("F3", "全校畢業學生升學就讀學校之學制別 - 實用技能學程(日)學生數");
                LookupValues.Add("F3-A3", "全校畢業學生升學就讀學校之學制別 - 實用技能學程(日)比率");
                LookupValues.Add("G3", "全校畢業學生升學就讀學校之學制別 - 實用技能學程(夜)學生數");
                LookupValues.Add("G3-A3", "全校畢業學生升學就讀學校之學制別 - 實用技能學程(夜)比率");
                LookupValues.Add("H3", "全校畢業學生升學就讀學校之學制別 - 進修學校學生數");
                LookupValues.Add("H3-A3", "全校畢業學生升學就讀學校之學制別 - 進修學校比率");
                LookupValues.Add("I3", "全校畢業學生升學就讀學校之學制別 - 五專學生數");
                LookupValues.Add("I3-A3", "全校畢業學生升學就讀學校之學制別 - 五專比率");
                LookupValues.Add("J3", "全校畢業學生升學就讀學校之學制別 - 其他學生數");
                LookupValues.Add("J3-A3", "全校畢業學生升學就讀學校之學制別 - 其他比率");

                LookupValues.Add("A4", "全校畢業學生升學之入學方式情形 - 升學學生總數");

                LookupValues.Add("B4", "全校畢業學生升學之入學方式情形 - 免試入學-校內直升學生數");
                LookupValues.Add("B4-A4", "全校畢業學生升學之入學方式情形 - 免試入學-校內直升比率");

                LookupValues.Add("C4", "全校畢業學生升學之入學方式情形 - 免試入學-分區免試學生數");
                LookupValues.Add("C4-A4", "全校畢業學生升學之入學方式情形 - 免試入學-分區免試比率");

                LookupValues.Add("D4", "全校畢業學生升學之入學方式情形 - 免試入學-單獨招生學生數");
                LookupValues.Add("D4-A4", "全校畢業學生升學之入學方式情形 - 免試入學-單獨招生比率");

                LookupValues.Add("E4", "全校畢業學生升學之入學方式情形 - 免試入學-技優甄審學生數");
                LookupValues.Add("E4-A4", "全校畢業學生升學之入學方式情形 - 免試入學-技優甄審比率");

                LookupValues.Add("F4", "全校畢業學生升學之入學方式情形 - 特色招生-考試分發入學學生數");
                LookupValues.Add("F4-A4", "全校畢業學生升學之入學方式情形 - 特色招生-考試分發入學比率");

                LookupValues.Add("G4", "全校畢業學生升學之入學方式情形 - 特色招生-職業類群科學生數");
                LookupValues.Add("G4-A4", "全校畢業學生升學之入學方式情形 - 特色招生-職業類群科比率");

                LookupValues.Add("H4", "全校畢業學生升學之入學方式情形 - 特色招生-藝才班學生數");
                LookupValues.Add("H4-A4", "全校畢業學生升學之入學方式情形 - 特色招生-藝才班比率");

                LookupValues.Add("I4", "全校畢業學生升學之入學方式情形 - 特色招生-體育班學生數");
                LookupValues.Add("I4-A4", "全校畢業學生升學之入學方式情形 - 特色招生-體育班比率");

                LookupValues.Add("J4", "全校畢業學生升學之入學方式情形 - 特色招生-科學班學生數");
                LookupValues.Add("J4-A4", "全校畢業學生升學之入學方式情形 - 特色招生-科學班比率");

                LookupValues.Add("K4", "全校畢業學生升學之入學方式情形 - 私校單獨招生學生數");
                LookupValues.Add("K4-A4", "全校畢業學生升學之入學方式情形 - 私校單獨招生比率");

                LookupValues.Add("L4", "全校畢業學生升學之入學方式情形 - 運動績優學生數");
                LookupValues.Add("L4-A4", "全校畢業學生升學之入學方式情形 - 運動績優比率");

                LookupValues.Add("M4", "全校畢業學生升學之入學方式情形 - 實用技能學程學生數");
                LookupValues.Add("M4-A4", "全校畢業學生升學之入學方式情形 - 實用技能學程比率");

                LookupValues.Add("N4", "全校畢業學生升學之入學方式情形 - 產業特殊需求學生數");
                LookupValues.Add("N4-A4", "全校畢業學生升學之入學方式情形 - 產業特殊需求比率");

                LookupValues.Add("O4", "全校畢業學生升學之入學方式情形 - 建教合作班學生數");
                LookupValues.Add("O4-A4", "全校畢業學生升學之入學方式情形 - 建教合作班比率");

                LookupValues.Add("P4", "全校畢業學生升學之入學方式情形 - 身心障礙生適性輔導安置學生數");
                LookupValues.Add("P4-A4", "全校畢業學生升學之入學方式情形 - 身心障礙生適性輔導安置比率");

                LookupValues.Add("Q4", "全校畢業學生升學之入學方式情形 - 五專免試入學學生數");
                LookupValues.Add("Q4-A4", "全校畢業學生升學之入學方式情形 - 五專免試入學比率");

                LookupValues.Add("R4", "全校畢業學生升學之入學方式情形 - 五專特色招生考試分發入學學生數");
                LookupValues.Add("R4-A4", "全校畢業學生升學之入學方式情形 - 五專特色招生考試分發入學比率");

                LookupValues.Add("S4", "全校畢業學生升學之入學方式情形 - 其他學生數");
                LookupValues.Add("S4-A4", "全校畢業學生升學之入學方式情形 - 其他比率");

                ApproachStatistics AppSat = ApproachSats.Find(x => ("" + x.SurveyYear).Equals(Value));

                if (AppSat != null)
                {
                    XElement elmContent = XElement.Load(new StringReader(AppSat.Content));

                    foreach (string Name in LookupValues.Keys)
                    {
                        XElement elmValue = elmContent.Element(Name);

                        if (elmValue != null)
                            grdContent.Rows.Add(LookupValues[Name], elmValue.Value);
                    }
                }
            }
        }
    }
}
