using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Aspose.Cells;
using DesktopLib;
using FISCA;
using FISCA.Data;
using FISCA.DSA;

namespace KHJHCentralOffice
{
    internal class ExportQueryData
    {
        private static string SchoolName = "名稱";

        public void Export()
        {
            SQLForm sql = new SQLForm();
            if (sql.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string sqltext = sql.SQLText;
                ExecuteTasks(sqltext);
            }
        }

        private static void ExecuteTasks(string sql)
        {
            MultiTaskingRunner runner = new MultiTaskingRunner();

            List<XElement> rsps = new List<XElement>();

            foreach (string uid in Program.MainPanel.SelectedSource)
            {
                string name = Program.GlobalSchoolCache[uid].Title;
                ConnectionHelper ch = ConnectionHelper.GetConnection(uid);
                runner.AddTask(string.Format("{0}({1})", name, uid), (x) =>
                {
                    XElement req = new XElement("Request",
                        new XElement("SQL", sql));

                    Envelope rsp = ch.CallService("UDTService.DML.Query", new Envelope(new XStringHolder(req)));

                    lock (rsps)
                    {
                        XElement r = XElement.Parse(rsp.BodyContent.XmlString);
                        r.SetAttributeValue("UID", ((object[])x)[0].ToString());
                        rsps.Add(r);
                    }
                }, new object[] { uid, ch }, new CancellationTokenSource());
            }

            runner.ExecuteTasks();

            ExportToExcel(rsps);
        }

        private static void ExportToExcel(List<XElement> rsps)
        {
            try
            {
                if (rsps.Count <= 0)
                    return;

                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "*.xls|*.xls";

                string filename = string.Empty;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    filename = dialog.FileName;

                    Workbook book = new Workbook();
                    book.Worksheets.Clear();

                    Worksheet sheet = book.Worksheets[book.Worksheets.Add()];

                    Dictionary<string, int> FNtoExcelIndex = OutputColumnName(sheet, rsps[0]);
                    Dictionary<int, string> indexToFIeldName = GetIndexMapping(rsps[0]);

                    int row = 1;
                    foreach (XElement rsp in rsps)
                    {
                        string uid = rsp.AttributeText("UID");
                        string name = Program.GlobalSchoolCache[uid].Title;

                        foreach (XElement record in rsp.Elements("Record"))
                        {
                            sheet.Cells[row, 0].PutValue(name);
                            foreach (XElement column in record.Elements("Column"))
                            {
                                int index = int.Parse(column.AttributeText("Index"));
                                string value = (column.Value);

                                sheet.Cells[row, FNtoExcelIndex[indexToFIeldName[index]]].PutValue(value);
                            }

                            row++;
                        }
                    }

                    book.Save(filename);
                }

                Process.Start(filename);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static Dictionary<int, string> GetIndexMapping(XElement xElement)
        {
            XElement metadata = xElement.Element("Metadata");
            Dictionary<int, string> result = new Dictionary<int, string>();

            int offset = 0;
            foreach (XElement column in metadata.Elements("Column"))
            {
                string index = column.AttributeText("Index");
                string value = column.AttributeText("Field");

                result.Add(int.Parse(index), value);

                offset++;
            }

            return result;
        }

        private static Dictionary<string, int> OutputColumnName(Worksheet sheet, XElement xElement)
        {
            XElement metadata = xElement.Element("Metadata");
            Dictionary<string, int> result = new Dictionary<string, int>();

            sheet.Cells[0, 0].PutValue(SchoolName);
            result.Add(SchoolName, 0);

            int offset = 1;
            foreach (XElement column in metadata.Elements("Column"))
            {
                string value = column.AttributeText("Field");

                sheet.Cells[0, offset].PutValue(value);
                result.Add(value, offset);

                offset++;
            }

            return result;
        }
    }
}
