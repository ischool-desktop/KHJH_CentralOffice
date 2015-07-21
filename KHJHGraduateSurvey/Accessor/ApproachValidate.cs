using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Aspose.Words;
using EMBA.Validator;
using FISCA.LogAgent;
using FISCA.UDT;

namespace JH_KH_GraduateSurvey.Accessor
{
    public class ApproachValidate
    {
        /// <summary>
        /// 請傳入「填報年度」.
        /// </summary>
        public static Dictionary<int, List<MessageItem>> Execute(int Year, Dictionary<int, Dictionary<string, string>> Data)
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
            return Template_102.ProcessRequest(ShiftedYear, Data);
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

            public abstract Dictionary<int, List<MessageItem>> ProcessRequest(int Year, Dictionary<int, Dictionary<string, string>> Data);
        }

        /// <summary>
        /// The 'ConcreteHandler' class
        /// </summary>
        private class Template_102 : Template
        {
            public override Dictionary<int, List<MessageItem>> ProcessRequest(int Year, Dictionary<int, Dictionary<string, string>> Data)
            {
                if (Year != 102)
                {
                    return this.successor.ProcessRequest(Year, Data);
                }
                else
                {
                    Dictionary<int, List<MessageItem>> Messages = new Dictionary<int, List<MessageItem>>();

                    Data.Keys.ToList().ForEach((y) =>
                    {
                        Dictionary<string, string> x = Data[y];
                        Messages.Add(y, new List<MessageItem>());

                        string q1_string = x["升學與就業情形"].Trim();
                        string q2_string = x["升學：就讀學校情形"].Trim();
                        string q3_string = x["升學：學制別"].Trim();
                        string q4_string = x["升學：入學方式"].Trim();
                        string q5_string = x["未升學未就業：動向"].Trim();
                        string q6_string = x["是否需要教育部協助"].Trim();
                        string memo = x["備註"].Trim();

                        int q1_int;
                        int q2_int;
                        int q3_int;
                        int q4_int;
                        int q5_int;

                        #region  檢查「升學與就業情形」
                        if (!string.IsNullOrWhiteSpace(q1_string))
                        {
                            if (!int.TryParse(q1_string, out q1_int))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」限填正整數。"));
                            else
                            {
                                if (q1_int > 3 || q1_int < 1)
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」僅可填 1~3。"));
                            }
                        }
                        if (int.TryParse(q1_string, out q1_int))
                        {
                            //有填寫並且為 1
                            if (q1_int == 1)
                            {
                                //  升學：就讀學校情形，必須填寫且值為 1~8
                                if (!(int.TryParse(q2_string, out q2_int) && (q2_int > 0 && q2_int < 9)))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 1 時，「升學：就讀學校情形」必須填寫 1~8。"));
                                }
                                //  升學：學制別，必須填寫且值為 1~9
                                if (!(int.TryParse(q3_string, out q3_int) && (q3_int > 0 && q3_int < 10)))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 1 時，「升學：學制別」必須填寫 1~9。"));
                                }
                                //  升學：入學方式，必須填寫且值為 1~18
                                if (!(int.TryParse(q4_string, out q4_int) && (q4_int > 0 && q4_int < 19)))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 1 時，「升學：入學方式」必須填寫 1~18。"));
                                }
                            }
                            //  升學 或 就業
                            if (q1_int == 1 || q1_int == 2)
                            {
                                //  未升學未就業：動向，不得填寫
                                if (!string.IsNullOrEmpty(q5_string))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 1~2 時，「未升學未就業：動向」不得填寫。"));
                                }
                            }
                            if (q1_int == 3)
                            {
                                //  未升學未就業：動向，必須填寫且值為 1~6
                                if (!(int.TryParse(q5_string, out q5_int) && (q5_int > 0 && q5_int < 7)))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 3 時，「未升學未就業：動向」必須填寫 1~6。"));
                                }
                            }
                            //  非升學
                            if (q1_int == 3 || q1_int == 2)
                            {
                                //  升學：就讀學校情形，不得填寫
                                if (!string.IsNullOrEmpty(q2_string))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 2~3 時，「升學：就讀學校情形」不得填寫。"));
                                }
                                //  升學：入學方式，不得填寫
                                if (!string.IsNullOrEmpty(q3_string))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 2~3 時，「升學：入學方式」不得填寫。"));
                                }
                                //  升學：學制別，不得填寫
                                if (!string.IsNullOrEmpty(q4_string))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 2~3 時，「升學：學制別」不得填寫。"));
                                }
                            }
                        }
                        #endregion

                        #region 檢查就讀學校
                        if (!string.IsNullOrWhiteSpace(q2_string))
                        {
                            if (!int.TryParse(q2_string, out q2_int))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」限填正整數。"));
                            else
                            {
                                if (q2_int > 8 || q2_int < 1)
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」僅可填 1~8。"));
                            }
                        }
                        if (int.TryParse(q2_string, out q2_int))
                        {
                            if (q2_int == 5)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int != 8)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 5 時，學制別僅可填 8。"));
                                }

                                if (int.TryParse(q4_string, out q4_int))
                                {
                                    List<int> Contents = new List<int>() { 16, 17 };

                                    if (!Contents.Contains(q4_int))
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 5 時，入學方式僅可填 16 或 17。"));
                                }
                            }

                            if (q2_int == 6)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int != 9)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 6 時，學制別僅可填 9。"));
                                }

                                if (int.TryParse(q4_string, out q4_int))
                                {
                                    if (q4_int != 18)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 6 時，入學方式僅可填 18。"));
                                }
                            }

                            if (q2_int == 7 || q2_int == 8)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int != 9)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 7 或 8 時，學制別僅可填 9。"));
                                }

                                if (int.TryParse(q4_string, out q4_int))
                                {
                                    if (q4_int != 18)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 7 或 8 時，入學方式僅可填 18。"));
                                }
                            }
                        }
                        #endregion

                        #region 檢查入學方式
                        if (!string.IsNullOrWhiteSpace(q4_string))
                        {
                            if (!int.TryParse(q4_string, out q4_int))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」限填正整數。"));
                            else
                            {
                                if (q4_int > 18 || q4_int < 1)
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」僅可填 1~18。"));
                            }
                        }
                        if (q4_string.Equals("16") || q4_string.Equals("17"))
                        {
                            if (!q2_string.Equals("5"))
                            {
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 16 或 17 時，「就讀學校」僅可填 5。"));
                            }
                            if (!q3_string.Equals("8"))
                            {
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 16 或 17 時，「學制別」僅可填 8。"));
                            }
                        }
                        if (q4_string.Equals("18"))
                        {
                            //	就讀學校情形
                            if (!(q2_string.Equals("6") || q2_string.Equals("7") || q2_string.Equals("8")))
                            {
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 18 時，「就讀學校」僅可填 6 或 7 或 8。"));
                            }
                            //	學制別
                            if (!q3_string.Equals("9"))
                            {
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 18 時，「學制別」僅可填 9。"));
                            }
                        }
                        if (int.TryParse(q4_string, out q4_int))
                        {
                            //1. 入學方式填12，學制別填5或6。
                            if (q4_int == 12)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    List<int> Contents = new List<int>() { 5, 6 };

                                    if (!Contents.Contains(q3_int))
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 12 時，「學制別」僅可填 5 或 6。"));
                                }
                            }

                            //2. 入學方式填14，學制別僅可填4。
                            if (q4_int == 14)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int != 4)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 14 時，「學制別」僅可填 4。"));
                                }
                            }

                            //3. 入學方式填6，學制別僅可填1。
                            if (q4_int == 6)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int != 1)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 6 時，「學制別」僅可填 1。"));
                                }
                            }
                            //4. 入學方式填9，學制別僅可填3。
                            if (q4_int == 9)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int != 3)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 9 時，「學制別」僅可填 3。"));
                                }
                            }
                        }
                        #endregion

                        #region 檢查學制別(q3)
                        if (!string.IsNullOrWhiteSpace(q3_string))
                        {
                            if (!int.TryParse(q3_string, out q3_int))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」限填正整數。"));
                            else
                            {
                                if (q3_int > 9 || q3_int < 1)
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」僅可填 1~9。"));
                            }

                            //	學制別為其他(9)，則就讀學校情形填6或7或8；入學方式填18
                            if (q3_string.Equals("9"))
                            {
                                //	就讀學校情形
                                if (!(q2_string.Equals("6") || q2_string.Equals("7") || q2_string.Equals("8")))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 9 時，就讀學校僅可填 6 或 7 或 8。"));
                                }
                                //	入學方式
                                if (!q4_string.Equals("18"))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 9 時，入學方式僅可填 18。"));
                                }
                            }
                            //	學制別為五專，則入學方式填16，17；就讀學校情形填5
                            if (q3_string.Equals("8"))
                            {
                                if (!(q4_string.Equals("16") || q4_string.Equals("17")))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 8 時，入學方式僅可填 16 或 17。"));
                                }
                                if (!q2_string.Equals("5"))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 8 時，就讀學校僅可填 5。"));
                                }
                            }
                            //	學制別為建教合作班，則入學方式填14
                            if (q3_string.Equals("4"))
                            {
                                if (!q4_string.Equals("14"))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 4 時，入學方式僅可填 14。"));
                                }
                            }
                            //	學制別為實用技能班5, 6，則入學方式填12
                            if (q3_string.Equals("5") || q3_string.Equals("6"))
                            {
                                if (!q4_string.Equals("12"))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 5 或 6 時，入學方式僅可填 12。"));
                                }
                            }
                        }

                        #endregion

                        #region 檢查未升學未就業動向
                        if (!string.IsNullOrWhiteSpace(q5_string))
                        {
                            if (!int.TryParse(q5_string, out q5_int))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業動向」限填正整數。"));
                            else
                            {
                                if (q5_int > 6 || q5_int < 1)
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業動向」僅可填 1~6。"));
                            }
                            if (!q1_string.Equals("3"))
                            {
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業動向」有填寫時，「升學與就業情形」僅可填 3。"));
                            }
                            if (!string.IsNullOrEmpty(q2_string) || !string.IsNullOrEmpty(q3_string) || !string.IsNullOrEmpty(q4_string))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業動向」有填寫時，「升學：就讀學校情形」、「升學：學制別」、「升學：入學方式」不可填寫。"));
                        }
                        if (int.TryParse(q5_string, out q5_int))
                        {
                            if (q5_int == 2)
                            {
                                #region 若為在家需填寫「是」「否」需教育部協助選項。
                                if (string.IsNullOrEmpty(q6_string))
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業：動向」填寫 2 時，請選填「是」或「否」需教育部協助選項。"));
                                #endregion

                                #region 若填寫「是」，需在「備註」欄填寫聯絡電話及通訊地址
                                if (q6_string.Equals("是") && string.IsNullOrEmpty(memo))
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業：動向」填寫 2 並需教育部協助時，請於「備註」欄填寫聯絡電話及通訊地址。"));
                                #endregion
                            }

                            #region 若為1失聯，請於「備註」欄中註明失聯原因
                            if (q5_int == 1)
                                if (string.IsNullOrEmpty(memo))
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業：動向」填寫 1 時，請於「備註」欄中註明失聯原因(如家長不知學生去向、電話空號等)。"));
                            #endregion

                            #region 若為6其他，請於「備註」欄中註明情況。
                            if (q5_int == 6)
                                if (string.IsNullOrEmpty(memo))
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業：動向」填寫 6 時，請於「備註」欄中註明情況。"));
                            #endregion
                        }
                        #endregion

                        #region 檢查需教育部協助
                        if (!string.IsNullOrEmpty(q6_string))
                        {
                            if (!(q6_string.Equals("是") || q6_string.Equals("否")))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「需教育部協助」僅能填寫「是」或「否」。"));
                        }
                        #endregion
                    });
                    return Messages;
                }
            }
        }

        private class Template_103 : Template
        {
            public override Dictionary<int, List<MessageItem>> ProcessRequest(int Year, Dictionary<int, Dictionary<string, string>> Data)
            {
                if (Year != 103)
                {
                    return this.successor.ProcessRequest(Year, Data);
                }
                else
                {
                    Dictionary<int, List<MessageItem>> Messages = new Dictionary<int, List<MessageItem>>();

                    Data.Keys.ToList().ForEach((y) =>
                    {
                        Dictionary<string, string> x = Data[y];
                        Messages.Add(y, new List<MessageItem>());

                        string q1_string = x["升學與就業情形"].Trim();
                        string q2_string = x["升學：就讀學校情形"].Trim();
                        string q3_string = x["升學：學制別"].Trim();
                        string q4_string = x["升學：入學方式"].Trim();
                        string q5_string = x["未升學未就業：動向"].Trim();
                        string q6_string = x["是否需要教育部協助"].Trim();
                        string memo = x["備註"].Trim();

                        int q1_int;
                        int q2_int;
                        int q3_int;
                        int q4_int;
                        int q5_int;

                        #region  檢查「升學與就業情形」
                        if (!string.IsNullOrWhiteSpace(q1_string))
                        {
                            if (!int.TryParse(q1_string, out q1_int))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」限填正整數。"));
                            else
                            {
                                if (q1_int > 3 || q1_int < 1)
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」僅可填 1~3。"));
                            }
                        }
                        if (int.TryParse(q1_string, out q1_int))
                        {
                            //有填寫並且為 1
                            if (q1_int == 1)
                            {
                                //  升學：就讀學校情形，必須填寫且值為 1~8
                                if (!(int.TryParse(q2_string, out q2_int) && (q2_int > 0 && q2_int < 9)))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 1 時，「升學：就讀學校情形」必須填寫 1~8。"));
                                }
                                //  升學：學制別，必須填寫且值為 1~9
                                if (!(int.TryParse(q3_string, out q3_int) && (q3_int > 0 && q3_int < 10)))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 1 時，「升學：學制別」必須填寫 1~9。"));
                                }
                                //  升學：入學方式，必須填寫且值為 1~25
                                if (!(int.TryParse(q4_string, out q4_int) && (q4_int > 0 && q4_int < 26)))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 1 時，「升學：入學方式」必須填寫 1~25。"));
                                }
                            }
                            //  升學 或 就業
                            if (q1_int == 1 || q1_int == 2)
                            {
                                //  未升學未就業：動向，不得填寫
                                if (!string.IsNullOrEmpty(q5_string))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 1~2 時，「未升學未就業：動向」不得填寫。"));
                                }
                            }
                            if (q1_int == 3)
                            {
                                //  未升學未就業：動向，必須填寫且值為 1~11
                                if (!(int.TryParse(q5_string, out q5_int) && (q5_int > 0 && q5_int < 12)))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 3 時，「未升學未就業：動向」必須填寫 1~11。"));
                                }
                            }
                            //  非升學
                            if (q1_int == 3 || q1_int == 2)
                            {
                                //  升學：就讀學校情形，不得填寫
                                if (!string.IsNullOrEmpty(q2_string))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 2~3 時，「升學：就讀學校情形」不得填寫。"));
                                }
                                //  升學：入學方式，不得填寫
                                if (!string.IsNullOrEmpty(q3_string))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 2~3 時，「升學：入學方式」不得填寫。"));
                                }
                                //  升學：學制別，不得填寫
                                if (!string.IsNullOrEmpty(q4_string))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「升學與就業情形」填寫 2~3 時，「升學：學制別」不得填寫。"));
                                }
                            }
                        }
                        #endregion

                        #region 檢查就讀學校
                        if (!string.IsNullOrWhiteSpace(q2_string))
                        {
                            if (!int.TryParse(q2_string, out q2_int))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」限填正整數。"));
                            else
                            {
                                if (q2_int > 8 || q2_int < 1)
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」僅可填 1~8。"));
                            }
                        }
                        if (int.TryParse(q2_string, out q2_int))
                        {
                            if (q2_int == 1)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int > 3 || q3_int < 2)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 1 時，學制別僅可填 2,3。"));
                                }
                            }

                            if (q2_int == 3 || q2_int == 4)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    IEnumerable<int> Contents = new List<int>() { 1, 2, 4, 5, 6, 7 };
                                    if (!Contents.Contains(q3_int))
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 3 或 4 時，學制別僅可填 1 或 2 或 4 或 5 或 6 或 7。"));
                                }
                            }

                            if (q2_int == 5)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int != 8)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 5 時，學制別僅可填 8。"));
                                }
                            }

                            if (q2_int == 6 || q2_int == 7 || q2_int == 8)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int != 9)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 6 或 7 或 8 時，學制別僅可填 9。"));
                                }

                                if (int.TryParse(q4_string, out q4_int))
                                {
                                    if (q4_int != 25)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」填寫 6 或 7 或 8 時，入學方式僅可填 25。"));
                                }
                            }
                        }
                        #endregion

                        #region 檢查入學方式
                        if (!string.IsNullOrWhiteSpace(q4_string))
                        {
                            if (!int.TryParse(q4_string, out q4_int))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」限填正整數。"));
                            else
                            {
                                if (q4_int > 25 || q4_int < 1)
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」僅可填 1~25。"));
                            }
                        }
                        if (int.TryParse(q4_string, out q4_int))
                        {
                            // 入學方式填1，學制別填1,2,3,7,8，就讀學校填1,2,3,4,5。
                            if (q4_int == 1)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    List<int> Contents = new List<int>() { 1,2,3,7,8 };

                                    if (!Contents.Contains(q3_int))
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 1 時，「學制別」僅可填 1,2,3,7,8。"));
                                }

                                if (int.TryParse(q2_string, out q2_int))
                                {
                                    List<int> Contents = new List<int>() { 1,2,3,4,5 };

                                    if (!Contents.Contains(q2_int))
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 1 時，「就讀學校」僅可填 1,2,3,4,5。"));
                                }
                            }
                            // 入學方式填9，學制別僅可7。
                            if (q4_int == 9)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int != 7)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 9 時，「學制別」僅可填 7。"));
                                }
                            }

                            // 入學方式填13，就讀學校僅可填1。
                            if (q4_int == 13)
                            {
                                if (int.TryParse(q2_string, out q2_int))
                                {
                                    if (q2_int != 1)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 13 時，「就讀學校」僅可填 1。"));
                                }
                            }

                            // 入學方式填17，學制別僅可填1。
                            if (q4_int == 17)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    if (q3_int != 1)
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 17 時，「學制別」僅可填 1。"));
                                }
                            }

                            // 入學方式填18，就讀學校僅可填 2 或 4。
                            if (q4_int == 18)
                            {
                                if (int.TryParse(q2_string, out q2_int))
                                {
                                    List<int> Contents = new List<int>() { 2, 4 };

                                    if (!Contents.Contains(q2_int))
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 18 時，「就讀學校」僅可填 2 或 4。"));
                                }
                            }

                            // 入學方式填22，學制別僅可填5或6。
                            if (q4_int == 22)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    List<int> Contents = new List<int>() { 5, 6 };

                                    if (!Contents.Contains(q3_int))
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 22 時，「學制別」僅可填 5 或 6。"));
                                }
                            }

                            // 入學方式填24，學制別僅可填4。
                            if (q4_int == 24)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    List<int> Contents = new List<int>() { 4 };

                                    if (!Contents.Contains(q3_int))
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 24 時，「學制別」僅可填 4。"));
                                }
                            }

                            // 入學方式填25，學制別僅可填9，就讀學校僅可填6或7或8。
                            if (q4_int == 25)
                            {
                                if (int.TryParse(q3_string, out q3_int))
                                {
                                    List<int> Contents = new List<int>() { 9 };

                                    if (!Contents.Contains(q3_int))
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 25 時，「學制別」僅可填 9。"));
                                }

                                if (int.TryParse(q2_string, out q2_int))
                                {
                                    List<int> Contents = new List<int>() { 6, 7, 8 };

                                    if (!Contents.Contains(q2_int))
                                        Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「入學方式」填寫 25 時，「就讀學校」僅可填 6 或 7 或 8。"));
                                }
                            }

                        }
                        #endregion

                        #region 檢查學制別(q3)
                        if (!string.IsNullOrWhiteSpace(q3_string))
                        {
                            if (!int.TryParse(q3_string, out q3_int))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」限填正整數。"));
                            else
                            {
                                if (q3_int > 9 || q3_int < 1)
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「就讀學校」僅可填 1~9。"));
                            }

                            //	學制別為建教合作班，則入學方式填24
                            if (q3_string.Equals("4"))
                            {
                                if (!q4_string.Equals("24"))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 4 時，入學方式僅可填 24。"));
                                }
                            }
                            //	學制別為實用技能班5, 6，則入學方式填22
                            if (q3_string.Equals("5") || q3_string.Equals("6"))
                            {
                                if (!q4_string.Equals("22"))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 5 或 6 時，入學方式僅可填 22。"));
                                }
                            }
                            
                            //	學制別為8五專，則入學方式填1；就讀學校情形填5                                                        
                            if (q3_string.Equals("8"))
                            {
                                if (!(q4_string.Equals("1")))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 8 時，入學方式僅可填 1。"));
                                }
                                if (!q2_string.Equals("5"))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 8 時，就讀學校僅可填 5。"));
                                }
                            }
                            //	學制別為其他(9)，則就讀學校情形填6或7或8；入學方式填25
                            if (q3_string.Equals("9"))
                            {
                                //	就讀學校情形
                                if (!(q2_string.Equals("6") || q2_string.Equals("7") || q2_string.Equals("8")))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 9 時，就讀學校僅可填 6 或 7 或 8。"));
                                }
                                //	入學方式
                                if (!q4_string.Equals("25"))
                                {
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「學制別」填寫 9 時，入學方式僅可填 25。"));
                                }
                            }
                        }

                        #endregion

                        #region 檢查未升學未就業動向
                        if (!string.IsNullOrWhiteSpace(q5_string))
                        {
                            if (!int.TryParse(q5_string, out q5_int))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業動向」限填正整數。"));
                            else
                            {
                                if (q5_int > 11 || q5_int < 1)
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業動向」僅可填 1~11。"));
                            }
                            if (!q1_string.Equals("3"))
                            {
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業動向」有填寫時，「升學與就業情形」僅可填 3。"));
                            }
                            if (!string.IsNullOrEmpty(q2_string) || !string.IsNullOrEmpty(q3_string) || !string.IsNullOrEmpty(q4_string))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業動向」有填寫時，「升學：就讀學校情形」、「升學：學制別」、「升學：入學方式」不可填寫。"));
                        }
                        if (int.TryParse(q5_string, out q5_int))
                        {
                            if (q5_int == 9)
                            {
                                #region 若為「尚未規劃」需填寫「是、1」「否、2」需教育部協助選項。
                                List<string> Contents = new List<string>() { "1", "2", "是", "否" };
                                if (!Contents.Contains(q6_string))
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業：動向」填寫 9 時，請選填「1(是)」或「2(否)」需教育部協助選項。"));
                                #endregion

                                #region 若填寫「是」，需在「備註」欄填寫聯絡電話及通訊地址
                                if ((q6_string.Equals("是") || q6_string.Equals("1")) && string.IsNullOrEmpty(memo))
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "填寫需教育部協助時，請於「備註」欄填寫居住地址、關係人連絡電話(手機或市內電話)。"));
                                #endregion
                            }

                            #region 若為失聯，請於「備註」欄中註明失聯原因
                            if (q5_int == 10)
                                if (string.IsNullOrEmpty(memo))
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業：動向」填寫 10 時，請於「備註」欄註記以下原因：電話更換或無人接聽/家長不知學生去向/離家或搬家/對方不願回應/其他(開放式填原因)。"));
                            #endregion

                            #region 若為其他動向，請於「備註」欄中註明情況。
                            if (q5_int == 11)
                                if (string.IsNullOrEmpty(memo))
                                    Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「未升學未就業：動向」填寫 11 時，請於「備註」欄中註明情況(開放式填答)。"));
                            #endregion
                        }
                        #endregion

                        #region 檢查需教育部協助
                        if (!string.IsNullOrEmpty(q6_string))
                        {
                            if (!(q6_string.Equals("是") || q6_string.Equals("否") || q6_string.Equals("1") || q6_string.Equals("2")))
                                Messages[y].Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "「需教育部協助」僅能填寫「1(是)」或「2(否)」。"));
                        }
                        #endregion
                    });
                    return Messages;
                }
            }
        }
    }
}
