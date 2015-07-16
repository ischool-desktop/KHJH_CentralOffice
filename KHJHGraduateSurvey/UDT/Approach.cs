using System;
using FISCA.UDT;

namespace JH_KH_GraduateSurvey.UDT
{
    [TableName("ischool.jh_kh.graduate_survey_approach")]
    public class Approach : ActiveRecord
    {
        /// <summary>
        /// 填報學年度
        /// </summary>
        [Field(Field = "survey_year", Indexed = true)]
        public int SurveyYear { get; set; }

        /// <summary>
        /// 學生系統編號
        /// </summary>
        [Field(Field = "ref_student_id", Indexed = true)]
        public int StudentID { get; set; }

        /// <summary>
        /// 升學與就業情形
        /// </summary>
        [Field(Field = "q1", Indexed = false)]
        public int Q1 { get; set; }

        /// <summary>
        /// 升學-就讀學校情形
        /// </summary>
        [Field(Field = "q2", Indexed = false)]
        public int? Q2 { get; set; }

        /// <summary>
        /// 升學-學制別
        /// </summary>
        [Field(Field = "q3", Indexed = false)]
        public int? Q3 { get; set; }

        /// <summary>
        /// 升學-入學方式
        /// </summary>
        [Field(Field = "q4", Indexed = false)]
        public int? Q4 { get; set; }

        /// <summary>
        /// 未升學未就業-動向
        /// </summary>
        [Field(Field = "q5", Indexed = false)]
        public int? Q5 { get; set; }

        /// <summary>
        /// 是否需要教育部協助
        /// </summary>
        [Field(Field = "q6", Indexed = false)]
        public string Q6 { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        [Field(Field = "memo", Indexed = false)]
        public string Memo { get; set; }

        /// <summary>
        /// 最後匯入時間
        /// </summary>
        [Field(Field = "last_update_time", Indexed = false)]
        public DateTime LastUpdateTime { get; set; }

        internal static void RaiseAfterUpdateEvent()
        {
            if (Approach.AfterUpdate != null)
                Approach.AfterUpdate(null, EventArgs.Empty);
        }

        internal static event EventHandler AfterUpdate;
    }
}
