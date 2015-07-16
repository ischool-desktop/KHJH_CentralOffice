using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace KHJHCentralOffice
{
    [TableName("ischool.jh_kh.open_time")]
    public class OpenTimeSetting : ActiveRecord
    {
        [Field(Field = "school_year")]
        public int SurveyYear { get; set; }

        [Field(Field = "start_date")]        
        public DateTime StartDate { get; set; }

        [Field(Field = "end_date")]
        public DateTime EndDate { get; set; }
    }
}