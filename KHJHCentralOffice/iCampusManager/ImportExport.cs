using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using FISCA;
using FISCA.Presentation;
using FISCA.UDT;

namespace KHJHCentralOffice
{
    internal class ImportExport
    {
        public ImportExport()
        {
            Program.MainPanel.RibbonBarItems["匯入/匯出"]["匯入"].Image = Properties.Resources.Import_Image;
            Program.MainPanel.RibbonBarItems["匯入/匯出"]["匯入"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            Program.MainPanel.RibbonBarItems["匯入/匯出"]["匯入"].Click += delegate
            {
                ReadSchoolRecords();
            };

            RibbonBarButton mb = Program.MainPanel.RibbonBarItems["匯入/匯出"]["匯出"];
            mb.Image = Properties.Resources.Export_Image;
            mb.Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            mb.Click += delegate
            {
                MessageBox.Show("製作中!");
            };
        }

        private void ReadSchoolRecords()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "*.xlsx|*.xlsx";
            dialog.Title = "選擇 DSNS 清單";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    MotherForm.SetStatusBarMessage("匯入中...");
                    Dictionary<string, School> schools = FromBackend();
                    ReadSchoolRecordFromExcel(dialog.FileName, schools);
                    schools.Values.SaveAll();
                    MessageBox.Show("匯入完成！");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    MotherForm.SetStatusBarMessage("");
                }
            }
        }

        private Dictionary<string, School> FromBackend()
        {
            AccessHelper access = new AccessHelper();
            List<School> lstSchools = access.Select<School>();
            return lstSchools.ToDictionary<School, string>(x => x.DSNS);
        }

        private void ReadSchoolRecordFromExcel(string fileName, Dictionary<string, School> schools)
        {
            try
            {
                Dictionary<string, int> column_map = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);

                Workbook book = new Workbook();
                book.Open(fileName, FileFormatType.Excel2007Xlsx);
                Worksheet sheet = book.Worksheets[0];
                ReadColumnName(column_map, sheet);

                for (int i = 1; i <= sheet.Cells.MaxDataRow; i++)
                {
                    string name = sheet.Cells[i, column_map["Name".ToLower()]].StringValue;

                    School school = new School();

                    if (schools.ContainsKey(name))
                        school = schools[name];
                    else
                        schools.Add(name, school);

                    school.DSNS = sheet.Cells[i, column_map["Name".ToLower()]].StringValue;
                    school.Title = sheet.Cells[i, column_map["Title".ToLower()]].StringValue;

                    if (column_map.ContainsKey("Group"))
                        school.Group = sheet.Cells[i, column_map["Group".ToLower()]].StringValue;

                    if (column_map.ContainsKey("Memo"))
                        school.Comment = sheet.Cells[i, column_map["Memo".ToLower()]].StringValue;
                }
            }
            catch (Exception ex)
            {
                ErrorBox.Show("讀取 Excel 錯誤！", ex);
            }
        }

        private static void ReadColumnName(Dictionary<string, int> column_map, Worksheet sheet)
        {
            for (int i = 0; i <= sheet.Cells.MaxDataColumn; i++)
            {
                string column_name = sheet.Cells[0, i].StringValue.ToLower();

                if (string.IsNullOrWhiteSpace(column_name))
                    continue;

                column_map.Add(column_name, i);
            }
        }
    }
}
