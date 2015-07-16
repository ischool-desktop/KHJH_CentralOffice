using FISCA.UDT;

namespace KHJHCentralOffice
{
    [TableName("school")]
    public class School : ActiveRecord
    {
        [Field(Field = "title", Indexed = true)]
        public string Title { get; set; }

        [Field(Field = "dsns", Indexed = true)]
        public string DSNS { get; set; }

        [Field(Field = "group")]
        public string Group { get; set; }

        [Field(Field = "comment")]
        public string Comment { get; set; }
    }
}