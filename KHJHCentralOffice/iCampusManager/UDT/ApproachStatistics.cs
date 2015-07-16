using FISCA.UDT;

namespace KHJHCentralOffice
{
    /// <summary>
    /// 國中畢業學生進路統計資料
    /// </summary>
    [TableName("ischool.jh_kh.graduate_survey_approach_statistics")]
    public class ApproachStatistics : ActiveRecord
    {
        /// <summary>
        /// 所屬學校系統編號
        /// </summary>
        [Field(Field="ref_school_id")]
        public int RefSchoolID { get; set; }

        /// <summary>
        /// 填報學年度
        /// </summary>
        [Field(Field="survey_year")]
        public int SurveyYear { get; set; }
        
        /// <summary>
        /// 內容
        /// </summary>
        [Field(Field="content")]
        public string Content { get; set; }
    }
}