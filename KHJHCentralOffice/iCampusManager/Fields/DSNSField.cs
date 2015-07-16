using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DesktopLib;

namespace KHJHCentralOffice
{
    internal class DSNSField : ListPaneFieldImproved
    {
        public DSNSField()
            : base("DSNS", "DSNS")
        {
        }

        protected override IEnumerable<ListPaneFieldImproved.Value> GetDataAsync(IEnumerable<string> primaryKeys)
        {
            string cmd = "select uid,dsns from $school where uid in (@@PrimaryKeys);";
            List<Value> results = new List<Value>();

            cmd = cmd.ReplaceList("@@PrimaryKeys", primaryKeys);
            Backend.Select(cmd).Each(row =>
            {
                string id = row["uid"] + "";
                string title = row["dsns"] + "";
                results.Add(new Value(id, title, string.Empty));
            });

            return results;
        }
    }
}
