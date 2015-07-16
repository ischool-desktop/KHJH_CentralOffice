using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using KHJHCentralOffice;

namespace KHJHCentralOffice.Accessor
{
    public class ApproachReport
    {
        /// <summary>
        /// 請傳入「填報年度」.
        /// </summary>
        public static Task<Dictionary<string, object>> Execute(int Year)
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

            public abstract Task<Dictionary<string, object>> ProcessRequest(int Year);
        }

        /// <summary>
        /// The 'ConcreteHandler' class
        /// </summary>
        private class Statistics_102 : Statistics
        {
            public override Task<Dictionary<string, object>> ProcessRequest(int Year)
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

                    Task<Dictionary<string, object>> task = Task<Dictionary<string, object>>.Factory.StartNew(() =>
                    {
                        Dictionary<string, object> mergeKeyValue = new Dictionary<string, object>();
                        List<ApproachStatistics> Records = Utility.AccessHelper
                            .Select<ApproachStatistics>("survey_year=" + survey_year);

                        List<School> Schools = Utility.AccessHelper.Select<School>();

                        if (Records.Count == 0)
                            throw new Exception("本年度無填報資料。");

                        decimal A1_sum = 0; decimal B1_sum = 0; decimal C1_sum = 0; decimal D1_sum = 0;

                        decimal A2_sum = 0; decimal B2_sum = 0; decimal C2_sum = 0; decimal D2_sum = 0;

                        decimal A3_sum = 0; decimal B3_sum = 0; decimal C3_sum = 0; decimal D3_sum = 0; decimal E3_sum = 0;
                        decimal F3_sum = 0; decimal G3_sum = 0; decimal H3_sum = 0; decimal I3_sum = 0;

                        decimal A4_sum = 0; decimal B4_sum = 0; decimal C4_sum = 0; decimal D4_sum = 0; decimal E4_sum = 0;
                        decimal F4_sum = 0; decimal G4_sum = 0; decimal H4_sum = 0; decimal I4_sum = 0; decimal J4_sum = 0;

                        decimal A5_sum = 0; decimal B5_sum = 0; decimal C5_sum = 0; decimal D5_sum = 0; decimal E5_sum = 0;
                        decimal F5_sum = 0; decimal G5_sum = 0; decimal H5_sum = 0; decimal I5_sum = 0; decimal J5_sum = 0;
                        decimal K5_sum = 0; decimal L5_sum = 0; decimal M5_sum = 0; decimal N5_sum = 0; decimal O5_sum = 0;
                        decimal P5_sum = 0; decimal Q5_sum = 0; decimal R5_sum = 0; decimal S5_sum = 0;

                        int S1 = 0; int S2 = 0; int S3 = 0;

                        #region 統計資料
                        foreach (ApproachStatistics record in Records)
                        {
                            XElement elmApproach = XElement.Load(new StringReader(record.Content));

                            School vSchool = Schools
                                .Find(x => x.UID.Equals("" + record.RefSchoolID));

                            A1_sum += int.Parse(elmApproach.Element("A1").Value);

                            if (vSchool != null && vSchool.Group != null)
                            {
                                if (vSchool.Group.IndexOf("公立") >= 0)
                                {
                                    B1_sum += int.Parse(elmApproach.Element("A1").Value);
                                    S1++;
                                }

                                if (vSchool.Group.IndexOf("私立") >= 0)
                                {
                                    C1_sum += int.Parse(elmApproach.Element("A1").Value);
                                    S2++;
                                }

                                if (vSchool.Group.IndexOf("附") >= 0)
                                {
                                    S3++;
                                    D1_sum += int.Parse(elmApproach.Element("A1").Value);
                                }
                            }

                            A2_sum += int.Parse(elmApproach.Element("A1").Value);
                            B2_sum += int.Parse(elmApproach.Element("B1").Value);
                            C2_sum += int.Parse(elmApproach.Element("C1").Value);
                            D2_sum += int.Parse(elmApproach.Element("D1").Value);

                            A3_sum += int.Parse(elmApproach.Element("A2").Value);
                            B3_sum += int.Parse(elmApproach.Element("B2").Value);
                            C3_sum += int.Parse(elmApproach.Element("C2").Value);
                            D3_sum += int.Parse(elmApproach.Element("D2").Value);
                            E3_sum += int.Parse(elmApproach.Element("E2").Value);
                            F3_sum += int.Parse(elmApproach.Element("F2").Value);
                            G3_sum += int.Parse(elmApproach.Element("G2").Value);
                            H3_sum += int.Parse(elmApproach.Element("H2").Value);
                            I3_sum += int.Parse(elmApproach.Element("I2").Value);

                            A4_sum += int.Parse(elmApproach.Element("A3").Value);
                            B4_sum += int.Parse(elmApproach.Element("B3").Value);
                            C4_sum += int.Parse(elmApproach.Element("C3").Value);
                            D4_sum += int.Parse(elmApproach.Element("D3").Value);
                            E4_sum += int.Parse(elmApproach.Element("E3").Value);
                            F4_sum += int.Parse(elmApproach.Element("F3").Value);
                            G4_sum += int.Parse(elmApproach.Element("G3").Value);
                            H4_sum += int.Parse(elmApproach.Element("H3").Value);
                            I4_sum += int.Parse(elmApproach.Element("I3").Value);
                            J4_sum += int.Parse(elmApproach.Element("J3").Value);

                            A5_sum += int.Parse(elmApproach.Element("A4").Value);
                            B5_sum += int.Parse(elmApproach.Element("B4").Value);
                            C5_sum += int.Parse(elmApproach.Element("C4").Value);
                            D5_sum += int.Parse(elmApproach.Element("D4").Value);
                            E5_sum += int.Parse(elmApproach.Element("E4").Value);
                            F5_sum += int.Parse(elmApproach.Element("F4").Value);
                            G5_sum += int.Parse(elmApproach.Element("G4").Value);
                            H5_sum += int.Parse(elmApproach.Element("H4").Value);
                            I5_sum += int.Parse(elmApproach.Element("I4").Value);
                            J5_sum += int.Parse(elmApproach.Element("J4").Value);
                            K5_sum += int.Parse(elmApproach.Element("K4").Value);
                            L5_sum += int.Parse(elmApproach.Element("L4").Value);
                            M5_sum += int.Parse(elmApproach.Element("M4").Value);
                            N5_sum += int.Parse(elmApproach.Element("N4").Value);
                            O5_sum += int.Parse(elmApproach.Element("O4").Value);
                            P5_sum += int.Parse(elmApproach.Element("P4").Value);
                            Q5_sum += int.Parse(elmApproach.Element("Q4").Value);
                            R5_sum += int.Parse(elmApproach.Element("R4").Value);
                            S5_sum += int.Parse(elmApproach.Element("S4").Value);
                        }
                        #endregion

                        #region 全縣(市)公私立國中畢業生
                        mergeKeyValue.Add("A1", A1_sum);
                        mergeKeyValue.Add("S1", S1);
                        mergeKeyValue.Add("S2", S2);
                        mergeKeyValue.Add("S3", S3);

                        mergeKeyValue.Add("B1", B1_sum);
                        mergeKeyValue.Add("C1", C1_sum);
                        mergeKeyValue.Add("D1", D1_sum);
                        mergeKeyValue.Add("B1/A1", A1_sum > 0 ? Math.Round(B1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C1/A1", A1_sum > 0 ? Math.Round(C1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D1/A1", A1_sum > 0 ? Math.Round(D1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 全縣市畢業學生升學與就業情形
                        mergeKeyValue.Add("A2", A2_sum);

                        mergeKeyValue.Add("B2", B2_sum);
                        mergeKeyValue.Add("C2", C2_sum);
                        mergeKeyValue.Add("D2", D2_sum);
                        mergeKeyValue.Add("B2/A2", A2_sum > 0 ? Math.Round(B2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C2/A2", A2_sum > 0 ? Math.Round(C2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D2/A2", A2_sum > 0 ? Math.Round(D2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 全縣市畢業學生升學之就讀學校情形
                        mergeKeyValue.Add("A3", A3_sum);

                        mergeKeyValue.Add("B3", B3_sum);
                        mergeKeyValue.Add("C3", C3_sum);
                        mergeKeyValue.Add("D3", D3_sum);
                        mergeKeyValue.Add("E3", E3_sum);
                        mergeKeyValue.Add("F3", F3_sum);
                        mergeKeyValue.Add("G3", G3_sum);
                        mergeKeyValue.Add("H3", H3_sum);
                        mergeKeyValue.Add("I3", I3_sum);

                        mergeKeyValue.Add("B3/A3", B3_sum > 0 ? Math.Round(B3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C3/A3", C3_sum > 0 ? Math.Round(C3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D3/A3", D3_sum > 0 ? Math.Round(D3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E3/A3", E3_sum > 0 ? Math.Round(E3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F3/A3", F3_sum > 0 ? Math.Round(F3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G3/A3", G3_sum > 0 ? Math.Round(G3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H3/A3", H3_sum > 0 ? Math.Round(H3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I3/A3", I3_sum > 0 ? Math.Round(I3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 全縣市畢業學生升學就讀學校之學制別
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

                        mergeKeyValue.Add("B4/A4", A4_sum > 0 ? Math.Round(B4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C4/A4", A4_sum > 0 ? Math.Round(C4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D4/A4", A4_sum > 0 ? Math.Round(D4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E4/A4", A4_sum > 0 ? Math.Round(E4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F4/A4", A4_sum > 0 ? Math.Round(F4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G4/A4", A4_sum > 0 ? Math.Round(G4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H4/A4", A4_sum > 0 ? Math.Round(H4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I4/A4", A4_sum > 0 ? Math.Round(I4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("J4/A4", A4_sum > 0 ? Math.Round(J4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 全縣市畢業學生升學之入學方式情形
                        mergeKeyValue.Add("A5", A5_sum);

                        mergeKeyValue.Add("B5", B5_sum);
                        mergeKeyValue.Add("C5", C5_sum);
                        mergeKeyValue.Add("D5", D5_sum);
                        mergeKeyValue.Add("E5", E5_sum);
                        mergeKeyValue.Add("F5", F5_sum);
                        mergeKeyValue.Add("G5", G5_sum);
                        mergeKeyValue.Add("H5", H5_sum);
                        mergeKeyValue.Add("I5", I5_sum);
                        mergeKeyValue.Add("J5", J5_sum);

                        mergeKeyValue.Add("K5", K5_sum);
                        mergeKeyValue.Add("L5", L5_sum);
                        mergeKeyValue.Add("M5", M5_sum);
                        mergeKeyValue.Add("N5", N5_sum);
                        mergeKeyValue.Add("O5", O5_sum);
                        mergeKeyValue.Add("P5", P5_sum);
                        mergeKeyValue.Add("Q5", Q5_sum);
                        mergeKeyValue.Add("R5", R5_sum);
                        mergeKeyValue.Add("S5", S5_sum);

                        mergeKeyValue.Add("B5/A5", A5_sum > 0 ? Math.Round(B5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C5/A5", A5_sum > 0 ? Math.Round(C5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D5/A5", A5_sum > 0 ? Math.Round(D5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E5/A5", A5_sum > 0 ? Math.Round(E5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F5/A5", A5_sum > 0 ? Math.Round(F5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G5/A5", A5_sum > 0 ? Math.Round(G5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H5/A5", A5_sum > 0 ? Math.Round(H5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I5/A5", A5_sum > 0 ? Math.Round(I5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("J5/A5", A5_sum > 0 ? Math.Round(J5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);

                        mergeKeyValue.Add("K5/A5", A5_sum > 0 ? Math.Round(K5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("L5/A5", A5_sum > 0 ? Math.Round(L5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("M5/A5", A5_sum > 0 ? Math.Round(M5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("N5/A5", A5_sum > 0 ? Math.Round(N5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("O5/A5", A5_sum > 0 ? Math.Round(O5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("P5/A5", A5_sum > 0 ? Math.Round(P5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("Q5/A5", A5_sum > 0 ? Math.Round(Q5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("R5/A5", A5_sum > 0 ? Math.Round(R5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("S5/A5", A5_sum > 0 ? Math.Round(S5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
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
                //  非本年度且有其他責任鏈
                if (Year != 103 && this.successor != null)
                {
                    return this.successor.ProcessRequest(Year);
                }
                //  是本年度或無其他責任鏈(沒有新的邏輯就用本年度的邏輯)
                else
                {
                    string survey_year = Year + "";

                    Task<Dictionary<string, object>> task = Task<Dictionary<string, object>>.Factory.StartNew(() =>
                    {
                        Dictionary<string, object> mergeKeyValue = new Dictionary<string, object>();
                        List<ApproachStatistics> Records = Utility.AccessHelper
                            .Select<ApproachStatistics>("survey_year=" + survey_year);

                        List<School> Schools = Utility.AccessHelper.Select<School>();

                        if (Records.Count == 0)
                            throw new Exception("本年度無填報資料。");

                        decimal A1_sum = 0; decimal B1_sum = 0; decimal C1_sum = 0; decimal D1_sum = 0;

                        decimal A2_sum = 0; decimal B2_sum = 0; decimal C2_sum = 0; decimal D2_sum = 0;

                        decimal A3_sum = 0; decimal B3_sum = 0; decimal C3_sum = 0; decimal D3_sum = 0; decimal E3_sum = 0;
                        decimal F3_sum = 0; decimal G3_sum = 0; decimal H3_sum = 0; decimal I3_sum = 0;

                        decimal A4_sum = 0; decimal B4_sum = 0; decimal C4_sum = 0; decimal D4_sum = 0; decimal E4_sum = 0;
                        decimal F4_sum = 0; decimal G4_sum = 0; decimal H4_sum = 0; decimal I4_sum = 0; decimal J4_sum = 0;

                        decimal A5_sum = 0; decimal B5_sum = 0; decimal C5_sum = 0; decimal D5_sum = 0; decimal E5_sum = 0;
                        decimal F5_sum = 0; decimal G5_sum = 0; decimal H5_sum = 0; decimal I5_sum = 0; decimal J5_sum = 0;
                        decimal K5_sum = 0; decimal L5_sum = 0; decimal M5_sum = 0; decimal N5_sum = 0; decimal O5_sum = 0;
                        decimal P5_sum = 0; decimal Q5_sum = 0; decimal R5_sum = 0; decimal S5_sum = 0; decimal T5_sum = 0;
                        decimal U5_sum = 0; decimal V5_sum = 0; decimal W5_sum = 0; decimal X5_sum = 0; decimal Y5_sum = 0;
                        decimal Z5_sum = 0;

                        int S1 = 0; int S2 = 0; int S3 = 0;

                        #region 統計資料
                        foreach (ApproachStatistics record in Records)
                        {
                            XElement elmApproach = XElement.Load(new StringReader(record.Content));

                            School vSchool = Schools
                                .Find(x => x.UID.Equals("" + record.RefSchoolID));

                            A1_sum += int.Parse(elmApproach.Element("A1").Value);

                            if (vSchool != null && vSchool.Group != null)
                            {
                                if (vSchool.Group.IndexOf("公立") >= 0)
                                {
                                    B1_sum += int.Parse(elmApproach.Element("A1").Value);
                                    S1++;
                                }

                                if (vSchool.Group.IndexOf("私立") >= 0)
                                {
                                    C1_sum += int.Parse(elmApproach.Element("A1").Value);
                                    S2++;
                                }

                                if (vSchool.Group.IndexOf("附") >= 0)
                                {
                                    S3++;
                                    D1_sum += int.Parse(elmApproach.Element("A1").Value);
                                }
                            }

                            A2_sum += int.Parse(elmApproach.Element("A1").Value);
                            B2_sum += int.Parse(elmApproach.Element("B1").Value);
                            C2_sum += int.Parse(elmApproach.Element("C1").Value);
                            D2_sum += int.Parse(elmApproach.Element("D1").Value);

                            A3_sum += int.Parse(elmApproach.Element("A2").Value);
                            B3_sum += int.Parse(elmApproach.Element("B2").Value);
                            C3_sum += int.Parse(elmApproach.Element("C2").Value);
                            D3_sum += int.Parse(elmApproach.Element("D2").Value);
                            E3_sum += int.Parse(elmApproach.Element("E2").Value);
                            F3_sum += int.Parse(elmApproach.Element("F2").Value);
                            G3_sum += int.Parse(elmApproach.Element("G2").Value);
                            H3_sum += int.Parse(elmApproach.Element("H2").Value);
                            I3_sum += int.Parse(elmApproach.Element("I2").Value);

                            A4_sum += int.Parse(elmApproach.Element("A3").Value);
                            B4_sum += int.Parse(elmApproach.Element("B3").Value);
                            C4_sum += int.Parse(elmApproach.Element("C3").Value);
                            D4_sum += int.Parse(elmApproach.Element("D3").Value);
                            E4_sum += int.Parse(elmApproach.Element("E3").Value);
                            F4_sum += int.Parse(elmApproach.Element("F3").Value);
                            G4_sum += int.Parse(elmApproach.Element("G3").Value);
                            H4_sum += int.Parse(elmApproach.Element("H3").Value);
                            I4_sum += int.Parse(elmApproach.Element("I3").Value);
                            J4_sum += int.Parse(elmApproach.Element("J3").Value);

                            A5_sum += int.Parse(elmApproach.Element("A4").Value);
                            B5_sum += int.Parse(elmApproach.Element("B4").Value);
                            C5_sum += int.Parse(elmApproach.Element("C4").Value);
                            D5_sum += int.Parse(elmApproach.Element("D4").Value);
                            E5_sum += int.Parse(elmApproach.Element("E4").Value);
                            F5_sum += int.Parse(elmApproach.Element("F4").Value);
                            G5_sum += int.Parse(elmApproach.Element("G4").Value);
                            H5_sum += int.Parse(elmApproach.Element("H4").Value);
                            I5_sum += int.Parse(elmApproach.Element("I4").Value);
                            J5_sum += int.Parse(elmApproach.Element("J4").Value);
                            K5_sum += int.Parse(elmApproach.Element("K4").Value);
                            L5_sum += int.Parse(elmApproach.Element("L4").Value);
                            M5_sum += int.Parse(elmApproach.Element("M4").Value);
                            N5_sum += int.Parse(elmApproach.Element("N4").Value);
                            O5_sum += int.Parse(elmApproach.Element("O4").Value);
                            P5_sum += int.Parse(elmApproach.Element("P4").Value);
                            Q5_sum += int.Parse(elmApproach.Element("Q4").Value);
                            R5_sum += int.Parse(elmApproach.Element("R4").Value);
                            S5_sum += int.Parse(elmApproach.Element("S4").Value);
                            T5_sum += int.Parse(elmApproach.Element("T4").Value);
                            U5_sum += int.Parse(elmApproach.Element("U4").Value);
                            V5_sum += int.Parse(elmApproach.Element("V4").Value);
                            W5_sum += int.Parse(elmApproach.Element("W4").Value);
                            X5_sum += int.Parse(elmApproach.Element("X4").Value);
                            Y5_sum += int.Parse(elmApproach.Element("Y4").Value);
                            Z5_sum += int.Parse(elmApproach.Element("Z4").Value);
                        }
                        #endregion

                        #region 全縣(市)公私立國中畢業生
                        mergeKeyValue.Add("A1", A1_sum);
                        mergeKeyValue.Add("S1", S1);
                        mergeKeyValue.Add("S2", S2);
                        mergeKeyValue.Add("S3", S3);

                        mergeKeyValue.Add("B1", B1_sum);
                        mergeKeyValue.Add("C1", C1_sum);
                        mergeKeyValue.Add("D1", D1_sum);
                        mergeKeyValue.Add("B1/A1", A1_sum > 0 ? Math.Round(B1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C1/A1", A1_sum > 0 ? Math.Round(C1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D1/A1", A1_sum > 0 ? Math.Round(D1_sum * 100 / A1_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 全縣市畢業學生升學與就業情形
                        mergeKeyValue.Add("A2", A2_sum);

                        mergeKeyValue.Add("B2", B2_sum);
                        mergeKeyValue.Add("C2", C2_sum);
                        mergeKeyValue.Add("D2", D2_sum);
                        mergeKeyValue.Add("B2/A2", A2_sum > 0 ? Math.Round(B2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C2/A2", A2_sum > 0 ? Math.Round(C2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D2/A2", A2_sum > 0 ? Math.Round(D2_sum * 100 / A2_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 全縣市畢業學生升學之就讀學校情形
                        mergeKeyValue.Add("A3", A3_sum);

                        mergeKeyValue.Add("B3", B3_sum);
                        mergeKeyValue.Add("C3", C3_sum);
                        mergeKeyValue.Add("D3", D3_sum);
                        mergeKeyValue.Add("E3", E3_sum);
                        mergeKeyValue.Add("F3", F3_sum);
                        mergeKeyValue.Add("G3", G3_sum);
                        mergeKeyValue.Add("H3", H3_sum);
                        mergeKeyValue.Add("I3", I3_sum);

                        mergeKeyValue.Add("B3/A3", B3_sum > 0 ? Math.Round(B3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C3/A3", C3_sum > 0 ? Math.Round(C3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D3/A3", D3_sum > 0 ? Math.Round(D3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E3/A3", E3_sum > 0 ? Math.Round(E3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F3/A3", F3_sum > 0 ? Math.Round(F3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G3/A3", G3_sum > 0 ? Math.Round(G3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H3/A3", H3_sum > 0 ? Math.Round(H3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I3/A3", I3_sum > 0 ? Math.Round(I3_sum * 100 / A3_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 全縣市畢業學生升學就讀學校之學制別
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

                        mergeKeyValue.Add("B4/A4", A4_sum > 0 ? Math.Round(B4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C4/A4", A4_sum > 0 ? Math.Round(C4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D4/A4", A4_sum > 0 ? Math.Round(D4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E4/A4", A4_sum > 0 ? Math.Round(E4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F4/A4", A4_sum > 0 ? Math.Round(F4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G4/A4", A4_sum > 0 ? Math.Round(G4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H4/A4", A4_sum > 0 ? Math.Round(H4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I4/A4", A4_sum > 0 ? Math.Round(I4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("J4/A4", A4_sum > 0 ? Math.Round(J4_sum * 100 / A4_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        #region 全縣市畢業學生升學之入學方式情形
                        mergeKeyValue.Add("A5", A5_sum);

                        mergeKeyValue.Add("B5", B5_sum);
                        mergeKeyValue.Add("C5", C5_sum);
                        mergeKeyValue.Add("D5", D5_sum);
                        mergeKeyValue.Add("E5", E5_sum);
                        mergeKeyValue.Add("F5", F5_sum);
                        mergeKeyValue.Add("G5", G5_sum);
                        mergeKeyValue.Add("H5", H5_sum);
                        mergeKeyValue.Add("I5", I5_sum);
                        mergeKeyValue.Add("J5", J5_sum);

                        mergeKeyValue.Add("K5", K5_sum);
                        mergeKeyValue.Add("L5", L5_sum);
                        mergeKeyValue.Add("M5", M5_sum);
                        mergeKeyValue.Add("N5", N5_sum);
                        mergeKeyValue.Add("O5", O5_sum);
                        mergeKeyValue.Add("P5", P5_sum);
                        mergeKeyValue.Add("Q5", Q5_sum);
                        mergeKeyValue.Add("R5", R5_sum);
                        mergeKeyValue.Add("S5", S5_sum);
                        mergeKeyValue.Add("T5", T5_sum);
                        mergeKeyValue.Add("U5", U5_sum);
                        mergeKeyValue.Add("V5", V5_sum);
                        mergeKeyValue.Add("W5", W5_sum);
                        mergeKeyValue.Add("X5", X5_sum);
                        mergeKeyValue.Add("Y5", Y5_sum);
                        mergeKeyValue.Add("Z5", Z5_sum);

                        mergeKeyValue.Add("B5/A5", A5_sum > 0 ? Math.Round(B5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("C5/A5", A5_sum > 0 ? Math.Round(C5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("D5/A5", A5_sum > 0 ? Math.Round(D5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("E5/A5", A5_sum > 0 ? Math.Round(E5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("F5/A5", A5_sum > 0 ? Math.Round(F5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("G5/A5", A5_sum > 0 ? Math.Round(G5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("H5/A5", A5_sum > 0 ? Math.Round(H5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("I5/A5", A5_sum > 0 ? Math.Round(I5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("J5/A5", A5_sum > 0 ? Math.Round(J5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);

                        mergeKeyValue.Add("K5/A5", A5_sum > 0 ? Math.Round(K5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("L5/A5", A5_sum > 0 ? Math.Round(L5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("M5/A5", A5_sum > 0 ? Math.Round(M5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("N5/A5", A5_sum > 0 ? Math.Round(N5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("O5/A5", A5_sum > 0 ? Math.Round(O5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("P5/A5", A5_sum > 0 ? Math.Round(P5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("Q5/A5", A5_sum > 0 ? Math.Round(Q5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("R5/A5", A5_sum > 0 ? Math.Round(R5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("S5/A5", A5_sum > 0 ? Math.Round(S5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("T5/A5", A5_sum > 0 ? Math.Round(T5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("U5/A5", A5_sum > 0 ? Math.Round(U5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("V5/A5", A5_sum > 0 ? Math.Round(V5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("W5/A5", A5_sum > 0 ? Math.Round(W5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("X5/A5", A5_sum > 0 ? Math.Round(X5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("Y5/A5", A5_sum > 0 ? Math.Round(Y5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        mergeKeyValue.Add("Z5/A5", A5_sum > 0 ? Math.Round(Z5_sum * 100 / A5_sum, 2, MidpointRounding.AwayFromZero) : 0);
                        #endregion

                        return mergeKeyValue;
                    });
                    return task;
                }
            }
        }
    }
}
