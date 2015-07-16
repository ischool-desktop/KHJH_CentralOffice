using FISCA.UDT;

namespace KHJHCentralOffice
{
    /// <summary>
    /// 國中畢業生未升學未就業學生動向統計資料
    /// </summary>
    [TableName("ischool.jh_kh.graduate_survey_vagrant_statistics")]
    public class VagrantStatistics : ActiveRecord
    {
        /// <summary>
        /// 所屬學校系統編號
        /// </summary>
        [Field(Field="ref_school_id")]
        public int RefSchoolID { get; set;}
        
        /// <summary>
        /// 填報學年度
        /// </summary>
        [Field(Field="survey_year")]
        public int SurveyYear { get; set;}

        /// <summary>
        /// 已就業
        /// </summary>
        [Field(Field="in_job")]
        public int InJob { get; set; }

        /// <summary>
        /// 已就學
        /// </summary>
        [Field(Field="in_school")]
        public int InSchool { get; set; }

        /// <summary>
        /// 準備升學
        /// </summary>
        [Field(Field="prepare_school")]
        public int PrepareSchool { get; set; }

        /// <summary>
        /// 準備或正在找工作
        /// </summary>
        [Field(Field="prepare_job")]
        public int PrepareJob { get; set; }

        /// <summary>
        /// 參加職訓
        /// </summary>
        [Field(Field="in_tranining")]
        public int InTraining { get; set; }

        /// <summary>
        /// 家務勞動
        /// </summary>
        [Field(Field="in_home")]
        public int InHome { get; set; }

        /// <summary>
        /// 尚未規劃
        /// </summary>
        [Field(Field="no_plan")]
        public int NoPlan { get; set; }

        /// <summary>
        /// 失聯
        /// </summary>
        [Field(Field="dis_appearance")]
        public int DisAppearance { get; set; }

        /// <summary>
        /// 其他
        /// </summary>
        [Field(Field = "other")]
        public int Other { get; set; }
    }
}