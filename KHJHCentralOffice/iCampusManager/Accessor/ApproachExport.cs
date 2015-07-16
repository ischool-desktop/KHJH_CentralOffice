using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Aspose.Cells;
using KHJHCentralOffice;

namespace KHJHCentralOffice.Accessor
{
    public class ApproachExport
    {
        /// <summary>
        /// 請傳入「填報年度」.
        /// </summary>
        public static Task<Workbook> Execute(int Year)
        {
            Statistics Report_102 = new Statistics_102();
            Statistics Report_103 = new Statistics_103();

            //  設定責任鏈之關連
            Report_102.SetSuccessor(Report_103);

            //  開始責任鏈之走訪並回傳結果
            return Report_102.ProcessRequest(Year);
        }

        /// <summary>
        /// The 'Handler' abstract class
        /// </summary>
        private abstract class Statistics
        {
            protected Statistics successor;

            public void SetSuccessor(Statistics successor)
            {
                this.successor = successor;
            }

            public abstract Task<Workbook> ProcessRequest(int Year);
        }

        /// <summary>
        /// The 'ConcreteHandler' class
        /// </summary>
        private class Statistics_102 : Statistics
        {
            public override Task<Workbook> ProcessRequest(int Year)
            {
                //  非本年度且有其他責任鏈
                if (Year != 102 && this.successor != null)
                {
                    return this.successor.ProcessRequest(Year);
                }
                //  是本年度或無其他責任鏈(沒有新的邏輯就用本年度的邏輯)
                else
                {
                    string survey_year = Year + "";

                    Task<Workbook> task = Task<Workbook>.Factory.StartNew(() =>
                    {
                        string SQL = string.Format(@"select school.uid as 學校系統編號, ""group"" as 分類, title as 名稱, 
xpath_string(table_approach.content, 'A1') as 畢業總人數
, xpath_string(table_approach.content, 'B1') as 升學人數
, xpath_string(table_approach.content, 'B1-A1') as 升學比率
, xpath_string(table_approach.content, 'C1') as 就業人數
, xpath_string(table_approach.content, 'C1-A1') as 就業比率
, xpath_string(table_approach.content, 'D1') as 未升學未就業人數
, xpath_string(table_approach.content, 'D1-A1') as 未升學未就業比率
, xpath_string(table_approach.content, 'A2') as ""升學總人數(就讀學校)""
, xpath_string(table_approach.content, 'B2') as 公立高中人數
, xpath_string(table_approach.content, 'B2-A2') as 公立高中比率
, xpath_string(table_approach.content, 'C2') as 私立高中人數
, xpath_string(table_approach.content, 'C2-A2') as 私立高中比率
, xpath_string(table_approach.content, 'D2') as 公立高職人數
, xpath_string(table_approach.content, 'D2-A2') as 公立高職比率
, xpath_string(table_approach.content, 'E2') as 私立高職人數
, xpath_string(table_approach.content, 'E2-A2') as 私立高職比率
, xpath_string(table_approach.content, 'F2') as ""五專(就讀學校)人數""
, xpath_string(table_approach.content, 'F2-A2') as ""五專(就讀學校)比率""
, xpath_string(table_approach.content, 'G2') as 軍事學校人數
, xpath_string(table_approach.content, 'G2-A2') as 軍事學校比率
, xpath_string(table_approach.content, 'H2') as 赴國外或大陸就學人數
, xpath_string(table_approach.content, 'H2-A2') as 赴國外或大陸就學比率
, xpath_string(table_approach.content, 'I2') as ""其他(就讀學校)人數""
, xpath_string(table_approach.content, 'I2-A2') as ""其他(就讀學校)比率""
, xpath_string(table_approach.content, 'A3') as ""升學總人數(學制別)""
, xpath_string(table_approach.content, 'B3') as 職業類科人數
, xpath_string(table_approach.content, 'B3-A3') as 職業類科比率
, xpath_string(table_approach.content, 'C3') as 綜合高中人數
, xpath_string(table_approach.content, 'C3-A3') as 綜合高中比率
, xpath_string(table_approach.content, 'D3') as 普通高中人數
, xpath_string(table_approach.content, 'D3-A3') as 普通高中比率
, xpath_string(table_approach.content, 'E3') as ""建教合作班(學制別)人數""
, xpath_string(table_approach.content, 'E3-A3') as ""建教合作班(學制別)比率""
, xpath_string(table_approach.content, 'F3') as ""實用技能學程(日)人數""
, xpath_string(table_approach.content, 'F3-A3') as ""實用技能學程(日)比率""
, xpath_string(table_approach.content, 'G3') as ""實用技能學程(夜)人數""
, xpath_string(table_approach.content, 'G3-A3') as ""實用技能學程(夜)比率""
, xpath_string(table_approach.content, 'H3') as 進修學校人數
, xpath_string(table_approach.content, 'H3-A3') as 進修學校比率
, xpath_string(table_approach.content, 'I3') as ""五專(學制別)人數""
, xpath_string(table_approach.content, 'I3-A3') as ""五專(學制別)比率""
, xpath_string(table_approach.content, 'J3') as ""其他(學制別)人數""
, xpath_string(table_approach.content, 'J3-A3') as ""其他(學制別)比率""
, xpath_string(table_approach.content, 'A4') as ""升學總人數(入學方式)""
, xpath_string(table_approach.content, 'B4') as ""免試入學-校內直升人數""
, xpath_string(table_approach.content, 'B4-A4') as ""免試入學-校內直升比率""
, xpath_string(table_approach.content, 'C4') as ""免試入學-分區免試人數""
, xpath_string(table_approach.content, 'C4-A4') as ""免試入學-分區免試比率""
, xpath_string(table_approach.content, 'D4') as ""免試入學-單獨招生人數""
, xpath_string(table_approach.content, 'D4-A4') as ""免試入學-單獨招生比率""
, xpath_string(table_approach.content, 'E4') as ""免試入學-技優甄審人數""
, xpath_string(table_approach.content, 'E4-A4') as ""免試入學-技優甄審比率""
, xpath_string(table_approach.content, 'F4') as ""特色招生-考試分發入學人數""
, xpath_string(table_approach.content, 'F4-A4') as ""特色招生-考試分發入學比率""
, xpath_string(table_approach.content, 'G4') as ""特色招生-職業類群科人數""
, xpath_string(table_approach.content, 'G4-A4') as ""特色招生-職業類群科比率""
, xpath_string(table_approach.content, 'H4') as ""特色招生-藝才班人數""
, xpath_string(table_approach.content, 'H4-A4') as ""特色招生-藝才班比率""
, xpath_string(table_approach.content, 'I4') as ""特色招生-體育班人數""
, xpath_string(table_approach.content, 'I4-A4') as ""特色招生-體育班比率""
, xpath_string(table_approach.content, 'J4') as ""特色招生-科學班人數""
, xpath_string(table_approach.content, 'J4-A4') as ""特色招生-科學班比率""
, xpath_string(table_approach.content, 'K4') as 私校單獨招生人數
, xpath_string(table_approach.content, 'K4-A4') as 私校單獨招生比率
, xpath_string(table_approach.content, 'L4') as 運動績優人數
, xpath_string(table_approach.content, 'L4-A4') as 運動績優比率
, xpath_string(table_approach.content, 'M4') as 實用技能學程人數
, xpath_string(table_approach.content, 'M4-A4') as 實用技能學程比率
, xpath_string(table_approach.content, 'N4') as 產業特殊需求人數
, xpath_string(table_approach.content, 'N4-A4') as 產業特殊需求比率
, xpath_string(table_approach.content, 'O4') as ""建教合作班(入學方式)人數""
, xpath_string(table_approach.content, 'O4-A4') as ""建教合作班(入學方式)比率""
, xpath_string(table_approach.content, 'P4') as 身心障礙生適性輔導安置人數
, xpath_string(table_approach.content, 'P4-A4') as 身心障礙生適性輔導安置比率
, xpath_string(table_approach.content, 'Q4') as 五專免試人學人數
, xpath_string(table_approach.content, 'Q4-A4') as 五專免試人學比率
, xpath_string(table_approach.content, 'R4') as 五專特色招生考試分發人數
, xpath_string(table_approach.content, 'R4-A4') as 五專特色招生考試分發比率
, xpath_string(table_approach.content, 'S4') as ""其他(入學方式)人數""
, xpath_string(table_approach.content, 'S4-A4') as ""其他(入學方式)比率""
from $school as school 
left join 
(select ref_school_id as school_id, content from $ischool.jh_kh.graduate_survey_approach_statistics as approach where approach.survey_year={0}) as table_approach
on table_approach.school_id=school.uid
order by school.""group"", table_approach.content", survey_year);

                        DataTable table = (new FISCA.Data.QueryHelper()).Select(SQL);
                        //int i=0;
                        //foreach(DataColumn column in table.Columns)
                        //{
                        //	i++;
                        //	if (i < 4)
                        //		continue;

                        //	table.Columns[i-1].DataType = Type.GetType("System.Int32");
                        //}
                        Workbook wb = new Workbook();
                        wb.Worksheets[0].Cells.ImportDataTable(table, true, "A1");
                        wb.Worksheets[0].AutoFitColumns();
                        return wb;
                    });
                    return task;
                }
            }
        }

        /// <summary>
        /// The 'ConcreteHandler' class
        /// </summary>
        private class Statistics_103 : Statistics
        {
            public override Task<Workbook> ProcessRequest(int Year)
            {
                //  非本年度且有其他責任鏈
                if (Year != 103 && this.successor != null)
                {
                    return this.successor.ProcessRequest(Year);
                }
                //  是本年度或無其他責任鏈(沒有新的邏輯就用本年度的邏輯)
                else
                {
                    string survey_year = Year + "";

                    Task<Workbook> task = Task<Workbook>.Factory.StartNew(() =>
                    {
                        string SQL = string.Format(@"select school.uid as 學校系統編號, ""group"" as 分類, title as 名稱, 
xpath_string(table_approach.content, 'A1') as ""畢業總人數""
, xpath_string(table_approach.content, 'B1') as ""升學人數""
, xpath_string(table_approach.content, 'B1-A1') as ""升學比率""
, xpath_string(table_approach.content, 'C1') as ""就業人數""
, xpath_string(table_approach.content, 'C1-A1') as ""就業比率""
, xpath_string(table_approach.content, 'D1') as ""未升學未就業人數""
, xpath_string(table_approach.content, 'D1-A1') as ""未升學未就業比率""
, xpath_string(table_approach.content, 'A2') as ""升學總人數(就讀學校)""
, xpath_string(table_approach.content, 'B2') as ""公立高中人數""
, xpath_string(table_approach.content, 'B2-A2') as ""公立高中比率""
, xpath_string(table_approach.content, 'C2') as ""私立高中人數""
, xpath_string(table_approach.content, 'C2-A2') as ""私立高中比率""
, xpath_string(table_approach.content, 'D2') as ""公立高職人數""
, xpath_string(table_approach.content, 'D2-A2') as ""公立高職比率""
, xpath_string(table_approach.content, 'E2') as ""私立高職人數""
, xpath_string(table_approach.content, 'E2-A2') as ""私立高職比率""
, xpath_string(table_approach.content, 'F2') as ""五專(就讀學校)人數""
, xpath_string(table_approach.content, 'F2-A2') as ""五專(就讀學校)比率""
, xpath_string(table_approach.content, 'G2') as ""軍事學校人數""
, xpath_string(table_approach.content, 'G2-A2') as ""軍事學校比率""
, xpath_string(table_approach.content, 'H2') as ""赴國外或大陸就學人數""
, xpath_string(table_approach.content, 'H2-A2') as ""赴國外或大陸就學比率""
, xpath_string(table_approach.content, 'I2') as ""其他(就讀學校)人數""
, xpath_string(table_approach.content, 'I2-A2') as ""其他(就讀學校)比率""
, xpath_string(table_approach.content, 'A3') as ""升學總人數(學制別)""
, xpath_string(table_approach.content, 'B3') as ""職業類科人數""
, xpath_string(table_approach.content, 'B3-A3') as ""職業類科比率""
, xpath_string(table_approach.content, 'C3') as ""綜合高中人數""
, xpath_string(table_approach.content, 'C3-A3') as ""綜合高中比率""
, xpath_string(table_approach.content, 'D3') as ""普通高中人數""
, xpath_string(table_approach.content, 'D3-A3') as ""普通高中比率""
, xpath_string(table_approach.content, 'E3') as ""建教合作班(學制別)人數""
, xpath_string(table_approach.content, 'E3-A3') as ""建教合作班(學制別)比率""
, xpath_string(table_approach.content, 'F3') as ""實用技能學程(日)人數""
, xpath_string(table_approach.content, 'F3-A3') as ""實用技能學程(日)比率""
, xpath_string(table_approach.content, 'G3') as ""實用技能學程(夜)人數""
, xpath_string(table_approach.content, 'G3-A3') as ""實用技能學程(夜)比率""
, xpath_string(table_approach.content, 'H3') as ""進修學校人數""
, xpath_string(table_approach.content, 'H3-A3') as ""進修學校比率""
, xpath_string(table_approach.content, 'I3') as ""五專(學制別)人數""
, xpath_string(table_approach.content, 'I3-A3') as ""五專(學制別)比率""
, xpath_string(table_approach.content, 'J3') as ""其他(學制別)人數""
, xpath_string(table_approach.content, 'J3-A3') as ""其他(學制別)比率""
, xpath_string(table_approach.content, 'A4') as ""升學總人數(入學方式)""
, xpath_string(table_approach.content, 'B4') as ""免試入學-分區免試人數""
, xpath_string(table_approach.content, 'B4-A4') as ""免試入學-分區免試比率""
, xpath_string(table_approach.content, 'C4') as ""免試入學-校內直升人數""
, xpath_string(table_approach.content, 'C4-A4') as ""免試入學-校內直升比率""
, xpath_string(table_approach.content, 'D4') as ""免試入學-技優甄審人數""
, xpath_string(table_approach.content, 'D4-A4') as ""免試入學-技優甄審比率""
, xpath_string(table_approach.content, 'E4') as ""免試入學-技術型及單科型人數""
, xpath_string(table_approach.content, 'E4-A4') as ""免試入學-技術型及單科型比率""
, xpath_string(table_approach.content, 'F4') as ""免試入學-園區生獨招人數""
, xpath_string(table_approach.content, 'F4-A4') as ""免試入學-園區生獨招比率""
, xpath_string(table_approach.content, 'G4') as ""免試入學-基北區產特人數""
, xpath_string(table_approach.content, 'G4-A4') as ""免試入學-基北區產特比率""
, xpath_string(table_approach.content, 'H4') as ""免試入學-宜蘭區專長生人數""
, xpath_string(table_approach.content, 'H4-A4') as ""免試入學-宜蘭區專長生比率""
, xpath_string(table_approach.content, 'I4') as ""免試入學-屏東區離島生人數""
, xpath_string(table_approach.content, 'I4-A4') as ""免試入學-屏東區離島生比率""
, xpath_string(table_approach.content, 'J4') as ""免試入學-進修學校非應屆獨招人數""
, xpath_string(table_approach.content, 'J4-A4') as ""免試入學-進修學校非應屆獨招比率""
, xpath_string(table_approach.content, 'K4') as ""免試入學-原住民藝能(實驗)班人數""
, xpath_string(table_approach.content, 'K4-A4') as ""免試入學-原住民藝能(實驗)班比率""
, xpath_string(table_approach.content, 'L4') as ""免試入學-台商子弟專案人數""
, xpath_string(table_approach.content, 'L4-A4') as ""免試入學-台商子弟專案比率""
, xpath_string(table_approach.content, 'M4') as ""特色招生-考試分發人數""
, xpath_string(table_approach.content, 'M4-A4') as ""特色招生-考試分發比率""
, xpath_string(table_approach.content, 'N4') as ""特色招生-科學班人數""
, xpath_string(table_approach.content, 'N4-A4') as ""特色招生-科學班比率""
, xpath_string(table_approach.content, 'O4') as ""特色招生-藝才班(競賽表現)人數""
, xpath_string(table_approach.content, 'O4-A4') as ""特色招生-藝才班(競賽表現)比率""
, xpath_string(table_approach.content, 'P4') as ""特色招生-藝才班(甄選入學)人數""
, xpath_string(table_approach.content, 'P4-A4') as ""特色招生-藝才班(甄選入學)比率""
, xpath_string(table_approach.content, 'Q4') as ""特色招生-體育班人數""
, xpath_string(table_approach.content, 'Q4-A4') as ""特色招生-體育班比率""
, xpath_string(table_approach.content, 'R4') as ""特色招生-職校特招人數""
, xpath_string(table_approach.content, 'R4-A4') as ""特色招生-職校特招比率""
, xpath_string(table_approach.content, 'S4') as ""私校單獨招生人數""
, xpath_string(table_approach.content, 'S4-A4') as ""私校單獨招生比率""
, xpath_string(table_approach.content, 'T4') as ""運動績優獨招人數""
, xpath_string(table_approach.content, 'T4-A4') as ""運動績優獨招比率""
, xpath_string(table_approach.content, 'U4') as ""運動績優甄試人數""
, xpath_string(table_approach.content, 'U4-A4') as ""運動績優甄試比率""
, xpath_string(table_approach.content, 'V4') as ""運動績優甄審人數""
, xpath_string(table_approach.content, 'V4-A4') as ""運動績優甄審比率""
, xpath_string(table_approach.content, 'W4') as ""實用技能學程人數""
, xpath_string(table_approach.content, 'W4-A4') as ""實用技能學程比率""
, xpath_string(table_approach.content, 'X4') as ""十二年就學安置人數""
, xpath_string(table_approach.content, 'X4-A4') as ""十二年就學安置比率""
, xpath_string(table_approach.content, 'Y4') as ""建教班人數""
, xpath_string(table_approach.content, 'Y4-A4') as ""建教班比率""
, xpath_string(table_approach.content, 'Z4') as ""其他人數""
, xpath_string(table_approach.content, 'Z4-A4') as ""其他比率""
from $school as school 
left join 
(select ref_school_id as school_id, content from $ischool.jh_kh.graduate_survey_approach_statistics as approach where approach.survey_year={0}) as table_approach
on table_approach.school_id=school.uid
order by school.""group"", table_approach.content", survey_year);

                        DataTable table = (new FISCA.Data.QueryHelper()).Select(SQL);
                        //int i=0;
                        //foreach(DataColumn column in table.Columns)
                        //{
                        //	i++;
                        //	if (i < 4)
                        //		continue;

                        //	table.Columns[i-1].DataType = Type.GetType("System.Int32");
                        //}
                        Workbook wb = new Workbook();
                        wb.Worksheets[0].Cells.ImportDataTable(table, true, "A1");
                        wb.Worksheets[0].AutoFitColumns();
                        return wb;
                    });
                    return task;
                }
            }
        }
    }
}
