using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Aspose.Words;
using FISCA.UDT;
using K12.Data;

namespace JH_KH_GraduateSurvey.Accessor
{
    public class ApproachStatistics
    {
        /// <summary>
        /// 請傳入「填報年度」.
        /// </summary>
        public static Task<Dictionary<string, object>> Execute(int Year)
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

            Statistics Report_102 = new Statistics_102();
            Statistics Report_103 = new Statistics_103();

            //  設定責任鏈之關連
            Report_102.SetSuccessor(Report_103);
            
            //  開始責任鏈之走訪並回傳結果
            return Report_102.ProcessRequest(ShiftedYear);
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

            public abstract Task<Dictionary<string, object>> ProcessRequest(int Year);
        }

        /// <summary>
        /// The 'ConcreteHandler' class
        /// </summary>
        private class Statistics_102 : Statistics
        {
            public override Task<Dictionary<string, object>> ProcessRequest(int Year)
            {
                if (Year != 102)
                {
                    return this.successor.ProcessRequest(Year);
                }
                else
                {
                    string survey_year = Year + "";

                    Task<Dictionary<string, object>> task = Task<Dictionary<string, object>>.Factory.StartNew(() =>
                    {
                        List<string> keys = new List<string>();
                        List<object> values = new List<object>();
                        Dictionary<string, object> mergeKeyValue = new Dictionary<string, object>();
                        List<UDT.Approach> Records = (new AccessHelper()).Select<UDT.Approach>("survey_year=" + survey_year);
                        if (Records.Count == 0)
                            throw new Exception("您所選擇的年度無填報資料。");

                        decimal A1_sum = Records.Count; decimal A2_sum = 0; decimal A3_sum = 0; decimal A4_sum = 0;
                        decimal B1_sum = 0; decimal B2_sum = 0; decimal B3_sum = 0; decimal B4_sum = 0;
                        decimal C1_sum = 0; decimal C2_sum = 0; decimal C3_sum = 0; decimal C4_sum = 0;
                        decimal D1_sum = 0; decimal D2_sum = 0; decimal D3_sum = 0; decimal D4_sum = 0;

                        decimal E2_sum = 0; decimal F2_sum = 0; decimal G2_sum = 0; decimal H2_sum = 0; decimal I2_sum = 0;
                        decimal E3_sum = 0; decimal F3_sum = 0; decimal G3_sum = 0; decimal H3_sum = 0; decimal I3_sum = 0; decimal J3_sum = 0;

                        decimal K4_sum = 0; decimal L4_sum = 0; decimal M4_sum = 0; decimal N4_sum = 0; decimal O4_sum = 0; decimal P4_sum = 0;
                        decimal Q4_sum = 0; decimal R4_sum = 0; decimal S4_sum = 0;

                        decimal E4_sum = 0; decimal F4_sum = 0; decimal G4_sum = 0; decimal H4_sum = 0; decimal I4_sum = 0; decimal J4_sum = 0;
                        foreach (UDT.Approach record in Records)
                        {
                            #region 升學或就業情形
                            if (record.Q1 == 1)
                                B1_sum += 1;
                            if (record.Q1 == 2)
                                C1_sum += 1;
                            if (record.Q1 == 3)
                                D1_sum += 1;
                            #endregion

                            #region 就讀學校
                            if (record.Q2 == 1)
                                B2_sum += 1;
                            if (record.Q2 == 2)
                                C2_sum += 1;
                            if (record.Q2 == 3)
                                D2_sum += 1;
                            if (record.Q2 == 4)
                                E2_sum += 1;
                            if (record.Q2 == 5)
                                F2_sum += 1;
                            if (record.Q2 == 6)
                                G2_sum += 1;
                            if (record.Q2 == 7)
                                H2_sum += 1;
                            if (record.Q2 == 8)
                                I2_sum += 1;
                            #endregion

                            #region 學制別
                            if (record.Q3 == 1)
                                B3_sum += 1;
                            if (record.Q3 == 2)
                                C3_sum += 1;
                            if (record.Q3 == 3)
                                D3_sum += 1;
                            if (record.Q3 == 4)
                                E3_sum += 1;
                            if (record.Q3 == 5)
                                F3_sum += 1;
                            if (record.Q3 == 6)
                                G3_sum += 1;
                            if (record.Q3 == 7)
                                H3_sum += 1;
                            if (record.Q3 == 8)
                                I3_sum += 1;
                            if (record.Q3 == 9)
                                J3_sum += 1;
                            #endregion

                            #region 入學方式
                            if (record.Q4 == 1)
                                B4_sum += 1;
                            if (record.Q4 == 2)
                                C4_sum += 1;
                            if (record.Q4 == 3)
                                D4_sum += 1;
                            if (record.Q4 == 4)
                                E4_sum += 1;
                            if (record.Q4 == 5)
                                F4_sum += 1;
                            if (record.Q4 == 6)
                                G4_sum += 1;
                            if (record.Q4 == 7)
                                H4_sum += 1;
                            if (record.Q4 == 8)
                                I4_sum += 1;
                            if (record.Q4 == 9)
                                J4_sum += 1;
                            if (record.Q4 == 10)
                                K4_sum += 1;
                            if (record.Q4 == 11)
                                L4_sum += 1;
                            if (record.Q4 == 12)
                                M4_sum += 1;
                            if (record.Q4 == 13)
                                N4_sum += 1;
                            if (record.Q4 == 14)
                                O4_sum += 1;
                            if (record.Q4 == 15)
                                P4_sum += 1;
                            if (record.Q4 == 16)
                                Q4_sum += 1;
                            if (record.Q4 == 17)
                                R4_sum += 1;
                            if (record.Q4 == 18)
                                S4_sum += 1;
                            #endregion

                        }

                        #region 全校畢業學生升學就業情形
                        mergeKeyValue.Add("A1", A1_sum);
                        mergeKeyValue.Add("B1", B1_sum);
                        mergeKeyValue.Add("C1", C1_sum);
                        mergeKeyValue.Add("D1", D1_sum);
                        mergeKeyValue.Add("B1/A1", Math.Round(B1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero));
                        mergeKeyValue.Add("C1/A1", Math.Round(C1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero));
                        mergeKeyValue.Add("D1/A1", Math.Round(D1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero));
                        #endregion

                        #region 全校畢業學生升學之就讀學校情形
                        A2_sum = B2_sum + C2_sum + D2_sum + E2_sum + F2_sum + G2_sum + H2_sum + I2_sum;
                        mergeKeyValue.Add("A2", A2_sum);
                        mergeKeyValue.Add("B2", B2_sum);
                        mergeKeyValue.Add("C2", C2_sum);
                        mergeKeyValue.Add("D2", D2_sum);
                        mergeKeyValue.Add("E2", E2_sum);
                        mergeKeyValue.Add("F2", F2_sum);
                        mergeKeyValue.Add("G2", G2_sum);
                        mergeKeyValue.Add("H2", H2_sum);
                        mergeKeyValue.Add("I2", I2_sum);
                        mergeKeyValue.Add("B2/A2", A2_sum > 0 ? Math.Round(B2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C2/A2", A2_sum > 0 ? Math.Round(C2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D2/A2", A2_sum > 0 ? Math.Round(D2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E2/A2", A2_sum > 0 ? Math.Round(E2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F2/A2", A2_sum > 0 ? Math.Round(F2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G2/A2", A2_sum > 0 ? Math.Round(G2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H2/A2", A2_sum > 0 ? Math.Round(H2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I2/A2", A2_sum > 0 ? Math.Round(I2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 學制別
                        A3_sum = B3_sum + C3_sum + D3_sum + E3_sum + F3_sum + G3_sum + H3_sum + I3_sum + J3_sum;
                        mergeKeyValue.Add("A3", A3_sum);
                        mergeKeyValue.Add("B3", B3_sum);
                        mergeKeyValue.Add("C3", C3_sum);
                        mergeKeyValue.Add("D3", D3_sum);
                        mergeKeyValue.Add("E3", E3_sum);
                        mergeKeyValue.Add("F3", F3_sum);
                        mergeKeyValue.Add("G3", G3_sum);
                        mergeKeyValue.Add("H3", H3_sum);
                        mergeKeyValue.Add("I3", I3_sum);
                        mergeKeyValue.Add("J3", J3_sum);
                        mergeKeyValue.Add("B3/A3", A3_sum > 0 ? Math.Round(B3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C3/A3", A3_sum > 0 ? Math.Round(C3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D3/A3", A3_sum > 0 ? Math.Round(D3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E3/A3", A3_sum > 0 ? Math.Round(E3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F3/A3", A3_sum > 0 ? Math.Round(F3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G3/A3", A3_sum > 0 ? Math.Round(G3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H3/A3", A3_sum > 0 ? Math.Round(H3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I3/A3", A3_sum > 0 ? Math.Round(I3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("J3/A3", A3_sum > 0 ? Math.Round(J3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 入學方式
                        A4_sum = B4_sum + C4_sum + D4_sum + E4_sum + F4_sum + G4_sum + H4_sum + I4_sum + J4_sum +
                                 K4_sum + L4_sum + M4_sum + N4_sum + O4_sum + P4_sum + Q4_sum + R4_sum + S4_sum;

                        mergeKeyValue.Add("A4", A4_sum);

                        mergeKeyValue.Add("B4", B4_sum);
                        mergeKeyValue.Add("C4", C4_sum);
                        mergeKeyValue.Add("D4", D4_sum);
                        mergeKeyValue.Add("E4", E4_sum);
                        mergeKeyValue.Add("F4", F4_sum);
                        mergeKeyValue.Add("G4", G4_sum);
                        mergeKeyValue.Add("H4", H4_sum);
                        mergeKeyValue.Add("I4", I4_sum);
                        mergeKeyValue.Add("J4", J4_sum);
                        mergeKeyValue.Add("K4", K4_sum);
                        mergeKeyValue.Add("L4", L4_sum);
                        mergeKeyValue.Add("M4", M4_sum);
                        mergeKeyValue.Add("N4", N4_sum);
                        mergeKeyValue.Add("O4", O4_sum);
                        mergeKeyValue.Add("P4", P4_sum);
                        mergeKeyValue.Add("Q4", Q4_sum);
                        mergeKeyValue.Add("R4", R4_sum);
                        mergeKeyValue.Add("S4", S4_sum);

                        mergeKeyValue.Add("B4/A4", A4_sum > 0 ? Math.Round(B4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C4/A4", A4_sum > 0 ? Math.Round(C4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D4/A4", A4_sum > 0 ? Math.Round(D4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E4/A4", A4_sum > 0 ? Math.Round(E4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F4/A4", A4_sum > 0 ? Math.Round(F4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G4/A4", A4_sum > 0 ? Math.Round(G4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H4/A4", A4_sum > 0 ? Math.Round(H4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I4/A4", A4_sum > 0 ? Math.Round(I4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("J4/A4", A4_sum > 0 ? Math.Round(J4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);

                        mergeKeyValue.Add("K4/A4", A4_sum > 0 ? Math.Round(K4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("L4/A4", A4_sum > 0 ? Math.Round(L4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("M4/A4", A4_sum > 0 ? Math.Round(M4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("N4/A4", A4_sum > 0 ? Math.Round(N4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("O4/A4", A4_sum > 0 ? Math.Round(O4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("P4/A4", A4_sum > 0 ? Math.Round(P4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("Q4/A4", A4_sum > 0 ? Math.Round(Q4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("R4/A4", A4_sum > 0 ? Math.Round(R4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("S4/A4", A4_sum > 0 ? Math.Round(S4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 學校代碼及名稱
                        mergeKeyValue.Add("學校代碼", K12.Data.School.Code);
                        mergeKeyValue.Add("填報學校", K12.Data.School.ChineseName);
                        mergeKeyValue.Add("填報學年度", survey_year);
                        mergeKeyValue.Add("DSNS", FISCA.Authentication.DSAServices.AccessPoint);
                        mergeKeyValue.Add("筆數", "" + Records.Count);
                        #endregion

                        #region 未升學未就業學生--上傳用

                        List<UDT.Approach> UnApproachs = Records
                            .FindAll(x => x.Q5.HasValue && x.Q5.Value >= 1 && x.Q5.Value <= 6);

                        if (UnApproachs.Count > 0)
                        {
                            List<StudentRecord> Students = Student.SelectByIDs(UnApproachs.Select(x => "" + x.StudentID));

                            XElement elmUnApproachs = new XElement("UnApproachStudents");

                            foreach (UDT.Approach UnApproach in UnApproachs)
                            {
                                XElement elmStudent = new XElement("Student");

                                StudentRecord vStudent = Students.Find(x => x.ID.Equals("" + UnApproach.StudentID));

                                if (vStudent != null)
                                {
                                    elmStudent.Add(new XElement("姓名", vStudent.Name));
                                    elmStudent.Add(new XElement("座號", K12.Data.Int.GetString(vStudent.SeatNo)));
                                    elmStudent.Add(new XElement("未升學未就業動向", K12.Data.Int.GetString(UnApproach.Q5)));
                                    elmStudent.Add(new XElement("是否需要教育部協助", UnApproach.Q6));
                                    elmStudent.Add(new XElement("備註", UnApproach.Memo));

                                    elmUnApproachs.Add(elmStudent);
                                }
                            }

                            mergeKeyValue.Add("UnApproachStudents", elmUnApproachs);
                        }
                        #endregion

                        return mergeKeyValue;
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
            public override Task<Dictionary<string, object>> ProcessRequest(int Year)
            {
                if (Year != 103)
                {
                    return this.successor.ProcessRequest(Year);
                }
                else
                {
                    string survey_year = Year + "";

                    Task<Dictionary<string, object>> task = Task<Dictionary<string, object>>.Factory.StartNew(() =>
                    {
                        List<string> keys = new List<string>();
                        List<object> values = new List<object>();
                        Dictionary<string, object> mergeKeyValue = new Dictionary<string, object>();
                        List<UDT.Approach> Records = (new AccessHelper()).Select<UDT.Approach>("survey_year=" + survey_year);
                        if (Records.Count == 0)
                            throw new Exception("您所選擇的年度無填報資料。");

                        decimal A1_sum = Records.Count; decimal A2_sum = 0; decimal A3_sum = 0; decimal A4_sum = 0;
                        decimal B1_sum = 0; decimal B2_sum = 0; decimal B3_sum = 0; decimal B4_sum = 0;
                        decimal C1_sum = 0; decimal C2_sum = 0; decimal C3_sum = 0; decimal C4_sum = 0;
                        decimal D1_sum = 0; decimal D2_sum = 0; decimal D3_sum = 0; decimal D4_sum = 0;

                        decimal E2_sum = 0; decimal F2_sum = 0; decimal G2_sum = 0; decimal H2_sum = 0; decimal I2_sum = 0;
                        decimal E3_sum = 0; decimal F3_sum = 0; decimal G3_sum = 0; decimal H3_sum = 0; decimal I3_sum = 0; decimal J3_sum = 0;

                        decimal K4_sum = 0; decimal L4_sum = 0; decimal M4_sum = 0; decimal N4_sum = 0; decimal O4_sum = 0; decimal P4_sum = 0;
                        decimal Q4_sum = 0; decimal R4_sum = 0; decimal S4_sum = 0; decimal T4_sum = 0; decimal U4_sum = 0; decimal V4_sum = 0;
                        decimal W4_sum = 0; decimal X4_sum = 0; decimal Y4_sum = 0; decimal Z4_sum = 0;

                        decimal E4_sum = 0; decimal F4_sum = 0; decimal G4_sum = 0; decimal H4_sum = 0; decimal I4_sum = 0; decimal J4_sum = 0;
                        foreach (UDT.Approach record in Records)
                        {
                            #region 升學或就業情形
                            if (record.Q1 == 1)
                                B1_sum += 1;
                            if (record.Q1 == 2)
                                C1_sum += 1;
                            if (record.Q1 == 3)
                                D1_sum += 1;
                            #endregion

                            #region 就讀學校
                            if (record.Q2 == 1)
                                B2_sum += 1;
                            if (record.Q2 == 2)
                                C2_sum += 1;
                            if (record.Q2 == 3)
                                D2_sum += 1;
                            if (record.Q2 == 4)
                                E2_sum += 1;
                            if (record.Q2 == 5)
                                F2_sum += 1;
                            if (record.Q2 == 6)
                                G2_sum += 1;
                            if (record.Q2 == 7)
                                H2_sum += 1;
                            if (record.Q2 == 8)
                                I2_sum += 1;
                            #endregion

                            #region 學制別
                            if (record.Q3 == 1)
                                B3_sum += 1;
                            if (record.Q3 == 2)
                                C3_sum += 1;
                            if (record.Q3 == 3)
                                D3_sum += 1;
                            if (record.Q3 == 4)
                                E3_sum += 1;
                            if (record.Q3 == 5)
                                F3_sum += 1;
                            if (record.Q3 == 6)
                                G3_sum += 1;
                            if (record.Q3 == 7)
                                H3_sum += 1;
                            if (record.Q3 == 8)
                                I3_sum += 1;
                            if (record.Q3 == 9)
                                J3_sum += 1;
                            #endregion

                            #region 入學方式
                            if (record.Q4 == 1)
                                B4_sum += 1;
                            if (record.Q4 == 2)
                                C4_sum += 1;
                            if (record.Q4 == 3)
                                D4_sum += 1;
                            if (record.Q4 == 4)
                                E4_sum += 1;
                            if (record.Q4 == 5)
                                F4_sum += 1;
                            if (record.Q4 == 6)
                                G4_sum += 1;
                            if (record.Q4 == 7)
                                H4_sum += 1;
                            if (record.Q4 == 8)
                                I4_sum += 1;
                            if (record.Q4 == 9)
                                J4_sum += 1;
                            if (record.Q4 == 10)
                                K4_sum += 1;
                            if (record.Q4 == 11)
                                L4_sum += 1;
                            if (record.Q4 == 12)
                                M4_sum += 1;
                            if (record.Q4 == 13)
                                N4_sum += 1;
                            if (record.Q4 == 14)
                                O4_sum += 1;
                            if (record.Q4 == 15)
                                P4_sum += 1;
                            if (record.Q4 == 16)
                                Q4_sum += 1;
                            if (record.Q4 == 17)
                                R4_sum += 1;
                            if (record.Q4 == 18)
                                S4_sum += 1;
                            if (record.Q4 == 19)
                                T4_sum += 1;
                            if (record.Q4 == 20)
                                U4_sum += 1;
                            if (record.Q4 == 21)
                                V4_sum += 1;
                            if (record.Q4 == 22)
                                W4_sum += 1;
                            if (record.Q4 == 23)
                                X4_sum += 1;
                            if (record.Q4 == 24)
                                Y4_sum += 1;
                            if (record.Q4 == 25)
                                Z4_sum += 1;
                            #endregion

                        }

                        #region 全校畢業學生升學就業情形
                        mergeKeyValue.Add("A1", A1_sum);
                        mergeKeyValue.Add("B1", B1_sum);
                        mergeKeyValue.Add("C1", C1_sum);
                        mergeKeyValue.Add("D1", D1_sum);
                        mergeKeyValue.Add("B1/A1", Math.Round(B1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero));
                        mergeKeyValue.Add("C1/A1", Math.Round(C1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero));
                        mergeKeyValue.Add("D1/A1", Math.Round(D1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero));
                        #endregion

                        #region 全校畢業學生升學之就讀學校情形
                        A2_sum = B2_sum + C2_sum + D2_sum + E2_sum + F2_sum + G2_sum + H2_sum + I2_sum;
                        mergeKeyValue.Add("A2", A2_sum);
                        mergeKeyValue.Add("B2", B2_sum);
                        mergeKeyValue.Add("C2", C2_sum);
                        mergeKeyValue.Add("D2", D2_sum);
                        mergeKeyValue.Add("E2", E2_sum);
                        mergeKeyValue.Add("F2", F2_sum);
                        mergeKeyValue.Add("G2", G2_sum);
                        mergeKeyValue.Add("H2", H2_sum);
                        mergeKeyValue.Add("I2", I2_sum);
                        mergeKeyValue.Add("B2/A2", A2_sum > 0 ? Math.Round(B2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C2/A2", A2_sum > 0 ? Math.Round(C2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D2/A2", A2_sum > 0 ? Math.Round(D2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E2/A2", A2_sum > 0 ? Math.Round(E2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F2/A2", A2_sum > 0 ? Math.Round(F2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G2/A2", A2_sum > 0 ? Math.Round(G2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H2/A2", A2_sum > 0 ? Math.Round(H2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I2/A2", A2_sum > 0 ? Math.Round(I2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 學制別
                        A3_sum = B3_sum + C3_sum + D3_sum + E3_sum + F3_sum + G3_sum + H3_sum + I3_sum + J3_sum;
                        mergeKeyValue.Add("A3", A3_sum);
                        mergeKeyValue.Add("B3", B3_sum);
                        mergeKeyValue.Add("C3", C3_sum);
                        mergeKeyValue.Add("D3", D3_sum);
                        mergeKeyValue.Add("E3", E3_sum);
                        mergeKeyValue.Add("F3", F3_sum);
                        mergeKeyValue.Add("G3", G3_sum);
                        mergeKeyValue.Add("H3", H3_sum);
                        mergeKeyValue.Add("I3", I3_sum);
                        mergeKeyValue.Add("J3", J3_sum);
                        mergeKeyValue.Add("B3/A3", A3_sum > 0 ? Math.Round(B3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C3/A3", A3_sum > 0 ? Math.Round(C3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D3/A3", A3_sum > 0 ? Math.Round(D3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E3/A3", A3_sum > 0 ? Math.Round(E3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F3/A3", A3_sum > 0 ? Math.Round(F3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G3/A3", A3_sum > 0 ? Math.Round(G3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H3/A3", A3_sum > 0 ? Math.Round(H3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I3/A3", A3_sum > 0 ? Math.Round(I3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("J3/A3", A3_sum > 0 ? Math.Round(J3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 入學方式
                        A4_sum = B4_sum + C4_sum + D4_sum + E4_sum + F4_sum + G4_sum + H4_sum + I4_sum + J4_sum +
                                 K4_sum + L4_sum + M4_sum + N4_sum + O4_sum + P4_sum + Q4_sum + R4_sum + S4_sum;

                        mergeKeyValue.Add("A4", A4_sum);

                        mergeKeyValue.Add("B4", B4_sum);
                        mergeKeyValue.Add("C4", C4_sum);
                        mergeKeyValue.Add("D4", D4_sum);
                        mergeKeyValue.Add("E4", E4_sum);
                        mergeKeyValue.Add("F4", F4_sum);
                        mergeKeyValue.Add("G4", G4_sum);
                        mergeKeyValue.Add("H4", H4_sum);
                        mergeKeyValue.Add("I4", I4_sum);
                        mergeKeyValue.Add("J4", J4_sum);
                        mergeKeyValue.Add("K4", K4_sum);
                        mergeKeyValue.Add("L4", L4_sum);
                        mergeKeyValue.Add("M4", M4_sum);
                        mergeKeyValue.Add("N4", N4_sum);
                        mergeKeyValue.Add("O4", O4_sum);
                        mergeKeyValue.Add("P4", P4_sum);
                        mergeKeyValue.Add("Q4", Q4_sum);
                        mergeKeyValue.Add("R4", R4_sum);
                        mergeKeyValue.Add("S4", S4_sum);
                        mergeKeyValue.Add("T4", T4_sum);
                        mergeKeyValue.Add("U4", U4_sum);
                        mergeKeyValue.Add("V4", V4_sum);
                        mergeKeyValue.Add("W4", W4_sum);
                        mergeKeyValue.Add("X4", X4_sum);
                        mergeKeyValue.Add("Y4", Y4_sum);
                        mergeKeyValue.Add("Z4", Z4_sum);

                        mergeKeyValue.Add("B4/A4", A4_sum > 0 ? Math.Round(B4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C4/A4", A4_sum > 0 ? Math.Round(C4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D4/A4", A4_sum > 0 ? Math.Round(D4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E4/A4", A4_sum > 0 ? Math.Round(E4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F4/A4", A4_sum > 0 ? Math.Round(F4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G4/A4", A4_sum > 0 ? Math.Round(G4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H4/A4", A4_sum > 0 ? Math.Round(H4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I4/A4", A4_sum > 0 ? Math.Round(I4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("J4/A4", A4_sum > 0 ? Math.Round(J4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);

                        mergeKeyValue.Add("K4/A4", A4_sum > 0 ? Math.Round(K4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("L4/A4", A4_sum > 0 ? Math.Round(L4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("M4/A4", A4_sum > 0 ? Math.Round(M4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("N4/A4", A4_sum > 0 ? Math.Round(N4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("O4/A4", A4_sum > 0 ? Math.Round(O4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("P4/A4", A4_sum > 0 ? Math.Round(P4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("Q4/A4", A4_sum > 0 ? Math.Round(Q4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("R4/A4", A4_sum > 0 ? Math.Round(R4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("S4/A4", A4_sum > 0 ? Math.Round(S4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("T4/A4", A4_sum > 0 ? Math.Round(T4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("U4/A4", A4_sum > 0 ? Math.Round(U4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("V4/A4", A4_sum > 0 ? Math.Round(V4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("W4/A4", A4_sum > 0 ? Math.Round(W4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("X4/A4", A4_sum > 0 ? Math.Round(X4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("Y4/A4", A4_sum > 0 ? Math.Round(Y4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("Z4/A4", A4_sum > 0 ? Math.Round(Z4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 學校代碼及名稱
                        mergeKeyValue.Add("學校代碼", K12.Data.School.Code);
                        mergeKeyValue.Add("填報學校", K12.Data.School.ChineseName);
                        mergeKeyValue.Add("填報學年度", survey_year);
                        mergeKeyValue.Add("DSNS", FISCA.Authentication.DSAServices.AccessPoint);
                        mergeKeyValue.Add("筆數", "" + Records.Count);
                        #endregion

                        #region 未升學未就業學生--上傳用

                        List<UDT.Approach> UnApproachs = Records
                            .FindAll(x => x.Q1 == 3);

                        if (UnApproachs.Count > 0)
                        {
                            List<StudentRecord> Students = Student.SelectByIDs(UnApproachs.Select(x => "" + x.StudentID));
                            List<ClassRecord> Classz = Class.SelectAll();

                            XElement elmUnApproachs = new XElement("UnApproachStudents");

                            foreach (UDT.Approach UnApproach in UnApproachs)
                            {
                                XElement elmStudent = new XElement("Student");

                                StudentRecord vStudent = Students.Find(x => x.ID.Equals("" + UnApproach.StudentID));

                                if (vStudent != null)
                                {
                                    elmStudent.Add(new XElement("班級", vStudent.Class.Name + ""));
                                    elmStudent.Add(new XElement("姓名", vStudent.Name + ""));
                                    elmStudent.Add(new XElement("座號", K12.Data.Int.GetString(vStudent.SeatNo)));
                                    elmStudent.Add(new XElement("學號", vStudent.StudentNumber + ""));
                                    elmStudent.Add(new XElement("未升學未就業動向", K12.Data.Int.GetString(UnApproach.Q5)));
                                    elmStudent.Add(new XElement("是否需要教育部協助", UnApproach.Q6));
                                    elmStudent.Add(new XElement("備註", UnApproach.Memo));

                                    elmUnApproachs.Add(elmStudent);
                                }
                            }

                            mergeKeyValue.Add("UnApproachStudents", elmUnApproachs);
                        }
                        #endregion

                        return mergeKeyValue;
                    });
                    return task;
                }
            }
        }
    }
}
