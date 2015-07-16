using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Words;
using FISCA.UDT;

namespace JH_KH_GraduateSurvey.Accessor
{
    public class ApproachComment
    {
        private static void AddComment(Worksheet worksheet, string note, byte columnIndex)
        {
            int commentIndex = worksheet.Comments.Add(0, columnIndex);
            Aspose.Cells.Comment comment = worksheet.Comments[commentIndex];
            comment.Note = note;
            comment.WidthCM = 20;
        }

        private static byte? GetColumnIndex(Worksheet worksheet, string columnName)
        {
            for (byte i = 0; i <= worksheet.Cells.MaxDataColumn; i++)
            {
                if ((worksheet.Cells[0, i].Value + "").Trim().ToLower() == columnName.Trim().ToLower())
                    return i;
            }
            return null;
        }
        /// <summary>
        /// 請傳入「填報年度」.
        /// </summary>
        public static void Execute(int Year, Worksheet worksheet)
        {
            int ShiftedYear = Year;
            List<int> Years = new List<int>();
            try
            {
                XDocument document = XDocument.Parse(Properties.Resources.Approach_Import, LoadOptions.None);
                IEnumerable<XElement> elements = document.Descendants("Schema").OrderBy(x => int.Parse(x.Attribute("SchoolYear").Value));
                foreach (XElement element in elements)
                {
                    int SurveyYear = int.Parse(element.Attribute("SchoolYear").Value);
                    Years.Add(SurveyYear);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            if (!Years.Contains(Year))
            {
                ShiftedYear = Years.Where(x => x < Year).Max();
            }

            Template Template_102 = new Template_102();
            Template Template_103 = new Template_103();

            //  設定責任鏈之關連
            Template_102.SetSuccessor(Template_103);

            //  開始責任鏈之走訪並回傳結果
            Template_102.ProcessRequest(ShiftedYear, worksheet);
        }

        /// <summary>
        /// The 'Handler' abstract class
        /// </summary>
        private abstract class Template
        {
            protected Template successor;

            public void SetSuccessor(Template successor)
            {
                this.successor = successor;
            }

            public abstract void ProcessRequest(int Year, Worksheet worksheet);
        }

        /// <summary>
        /// The 'ConcreteHandler' class
        /// </summary>
        private class Template_102 : Template
        {
            public override void ProcessRequest(int Year, Worksheet worksheet)
            {
                if (Year != 102)
                {
                    this.successor.ProcessRequest(Year, worksheet);
                }
                else
                {
                    string columnName = "升學與就業情形";
                    string note = "填代碼 1~3。(1：升學；2：就業；3：未升學未就業)\n填「1」者，請續填「升學：就讀學校情形」、「升學：入學方式」、「升學：學制別」。\n填「3」者，請續填「未升學未就業：動向」。";
                    byte? columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);

                    columnName = "升學：就讀學校情形";
                    note = "填代碼 1~8。(1：公立高中；2：私立高中；3：公立高職；4：私立高職；5：五專；6：軍事學校；7：赴國外或大陸就學；8：其他)";
                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);

                    columnName = "升學：入學方式";
                    note = "填代碼 1~18。(1：免試入學-校內直升；2：免試入學-分區免試；3：免試入學-單獨招生；4：免試入學-技優甄審；5：特色招生-考試分發入學；6：特色招生-職業類群科；7：特色招生-藝才班；8：特色招生-體育班；9：特色招生-科學班；10：私校單獨招生；11：運動績優；12：實用技能學程；13：產業特殊需求；14：建教合作班；15：身心障礙生適性輔導安置；16：五專免試入學；17：五專特色招生考試分發入學；18：其他)";

                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);
                    columnName = "升學：學制別";
                    note = "填代碼 1~9。(1：職業類科；2：綜合高中；3：普通高中；4：建教合作班；5：實用技能學程(日)；6：實用技能學程(夜)；7：進修學校；8：五專；9：其他)";

                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);

                    columnName = "未升學未就業：動向";
                    note = "填代碼 1~6。(1：失聯；2：在家；3：重考；4：出國；5：參加職訓；6：其他)";

                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);

                    columnName = "是否需要教育部協助";
                    note = "若未升學未就業：動向為2在家，請選填「是」或「否」需教育部協助選項，需教育部協助者請於「備註」欄填寫聯絡電話及通訊地址。";

                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);

                    columnName = "備註";
                    note = "需教育部協助者請於「備註」欄填寫聯絡電話及通訊地址、若未升學未就業：動向為1失聯，請於「備註」欄中註明失聯原因(如家長不知學生去向、電話空號等)、若未升學未就業：動向為6其他，請於「備註」欄中註明情況。";

                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);
                }
            }
        }


        /// <summary>
        /// The 'ConcreteHandler' class
        /// </summary>
        private class Template_103 : Template
        {
            public override void ProcessRequest(int Year, Worksheet worksheet)
            {
                if (Year != 103)
                {
                    this.successor.ProcessRequest(Year, worksheet);
                }
                else
                {
                    string columnName = "升學與就業情形";
                    string note = "填代碼 1~3。(1：升學；2：就業；3：未升學未就業)\n填「1」者，請續填「升學：就讀學校情形」、「升學：入學方式」、「升學：學制別」。\n填「3」者，請續填「未升學未就業：動向」。";
                    byte? columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);

                    columnName = "升學：就讀學校情形";
                    note = "填代碼 1~8。(1：公立高中；2：私立高中；3：公立高職；4：私立高職；5：五專；6：軍事學校；7：赴國外或大陸就學；8：其他)";
                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);

                    columnName = "升學：入學方式";
                    note = "填代碼 1~25。(1：免試入學-分區免試；2：免試入學-校內直升；3：免試入學-技優甄審；4：免試入學-技術型及單科型；5：免試入學-園區生獨招；6：免試入學-基北區產特；7：免試入學-宜蘭區專長生；8：免試入學-屏東區離島生；9：免試入學-進修學校非應屆獨招；10：免試入學-原住民藝能(實驗)班；11：免試入學-台商子弟專案；12：特色招生-考試分發；13：特色招生-科學班；14：特色招生-藝才班(競賽表現)；15：特色招生-藝才班(甄選入學)；16：特色招生-體育班；17：特色招生-職校特招；18：私校單獨招生；19：運體績優獨招；20：運體績優甄試；21：運體績優甄審；22：實用技能學程；23：十二年就學安置；24：建教班；25：其他)";

                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);
                    columnName = "升學：學制別";
                    note = "填代碼 1~9。(1：職業類科；2：綜合高中；3：普通高中；4：建教合作班；5：實用技能學程(日)；6：實用技能學程(夜)；7：進修學校；8：五專；9：其他)";

                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);

                    columnName = "未升學未就業：動向";
                    note = "填代碼 1~11。(1：已就業；2：已就學(於繳交截止日前如以就學或以就業需填之代號)；3：特教生；4：準備升學；5：準備或正在找工作；6：參加職訓；7：家務勞動；8：健康因素；9：尚未規劃；10：失聯；11：其他動向)";

                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);

                    columnName = "是否需要教育部協助";
                    note = "若「未升學未就業：動向」為「9：尚未規劃」，請選填「1(是)」或「2(否)」。";

                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);

                    columnName = "備註";
                    note = "填「需教育部協助」者，此欄位請填寫「居住地址、關係人連絡電話(手機或市內電話)；填「失聯」者，此欄位請註記以下原因「電話更換或無人接聽/家長不知學生去向/離家或搬家/對方不願回應/其他(開放式填原因)」；填「其他動向」者，此欄位開放式填答。";

                    columnIndex = GetColumnIndex(worksheet, columnName);
                    if (columnIndex.HasValue)
                        AddComment(worksheet, note, columnIndex.Value);
                }
            }
        }
    }
}
