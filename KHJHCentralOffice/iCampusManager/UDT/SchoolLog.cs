using FISCA.UDT;

namespace KHJHCentralOffice
{
    /// <summary>
    /// 學校歷程
    /// </summary>
    [TableName("school_log")]
    public class SchoolLog : FISCA.UDT.ActiveRecord
    {
        /// <summary>
        /// 學校端DSNS
        /// </summary>
        [Field(Field = "dsns")]
        public string DSNS { get; set;}

        /// <summary>
        /// 動作
        /// </summary>
        [Field(Field = "action")]
        public string Action { get; set; }

        /// <summary>
        /// 內容
        /// </summary>
        [Field(Field = "content")]
        public string Content { get; set; }

        /// <summary>
        /// 細節
        /// </summary>
        [Field(Field = "detail")]
        public string Detail { get; set; }

        /// <summary>
        /// 已讀
        /// </summary>
        [Field(Field = "read")]
        public bool Read { get; set; }

        /// <summary>
        /// 註解
        /// </summary>
        [Field(Field = "comment")]
        public string Comment { get; set; }
    }
}