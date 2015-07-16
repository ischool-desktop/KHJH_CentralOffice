using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Aspose.Words;
using FISCA.LogAgent;
using FISCA.UDT;

namespace JH_KH_GraduateSurvey.Accessor
{
    public class ApproachSave
    {
        /// <summary>
        /// 請傳入「填報年度」.
        /// </summary>
        public static string Execute(int Year, Dictionary<string, Dictionary<string, string>> Data)
        {
            Template Template_102 = new Template_102();

            //  設定責任鏈之關連
            Template_102.SetSuccessor(null);

            //  開始責任鏈之走訪並回傳結果
            return Template_102.ProcessRequest(Year, Data);
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

            public abstract string ProcessRequest(int Year, Dictionary<string, Dictionary<string, string>> Data);
        }

        /// <summary>
        /// The 'ConcreteHandler' class
        /// </summary>
        private class Template_102 : Template
        {
            public override string ProcessRequest(int Year, Dictionary<string, Dictionary<string, string>> Data)
            {
                if (Year != 102 && this.successor != null)
                {
                    return this.successor.ProcessRequest(Year, Data);
                }
                else
                {
                    StringBuilder strLog = new StringBuilder();
                    List<UDT.Approach> ExistingRecords = new List<UDT.Approach>();
                    AccessHelper Access = new AccessHelper();

                    strLog.AppendLine("詳細資料：");

                    ExistingRecords = Access.Select<UDT.Approach>();
                    //  要新增的 Record 
                    List<UDT.Approach> insertRecords = new List<UDT.Approach>();
                    //  要更新的 Record
                    List<UDT.Approach> updateRecords = new List<UDT.Approach>();
                    foreach (string key in Data.Keys)
                    {
                        //string id_number = Data[key]["身分證號"].Trim();
                        string q1_string = Data[key]["升學與就業情形"].Trim();
                        string q2_string = Data[key]["升學：就讀學校情形"].Trim();
                        string q3_string = Data[key]["升學：學制別"].Trim();
                        string q4_string = Data[key]["升學：入學方式"].Trim();
                        string q5_string = Data[key]["未升學未就業：動向"].Trim();
                        string q6_string = Data[key]["是否需要教育部協助"].Trim().Replace("1", "是").Replace("2", "否");
                        string memo = Data[key]["備註"].Trim();

                        strLog.AppendLine("填報學年度「" + Year + "」升學與就業情形「" + q1_string + "」升學：就讀學校情形「" + q2_string + "」升學：學制別「" + q3_string + "」升學：入學方式「" + q4_string + "」未升學未就業：動向「" + q5_string + "」是否需要教育部協助「" + q6_string + "」備註「" + memo + "」");

                        int q2_int;
                        int q3_int;
                        int q4_int;
                        int q5_int;

                        int student_id = int.Parse(key);

                        UDT.Approach record = new UDT.Approach();
                        IEnumerable<UDT.Approach> filterRecords = new List<UDT.Approach>();
                        filterRecords = ExistingRecords.Where(x => x.StudentID == student_id);

                        if (filterRecords.Count() > 0)
                        {
                            record = filterRecords.OrderByDescending(x => x.LastUpdateTime).ElementAt(0);
                            updateRecords.Add(record);
                        }
                        else
                        {
                            insertRecords.Add(record);
                        }

                        int q1 = int.Parse(q1_string);

                        record.StudentID = student_id;
                        record.SurveyYear = Year;
                        record.Q1 = q1;

                        if (int.TryParse(q2_string, out q2_int))
                            record.Q2 = q2_int;
                        else
                            record.Q2 = null;

                        if (int.TryParse(q3_string, out q3_int))
                            record.Q3 = q3_int;
                        else
                            record.Q3 = null;

                        if (int.TryParse(q4_string, out q4_int))
                            record.Q4 = q4_int;
                        else
                            record.Q4 = null;

                        if (int.TryParse(q5_string, out q5_int))
                            record.Q5 = q5_int;
                        else
                            record.Q5 = null;

                        record.Q6 = q6_string;
                        record.Memo = memo;

                        record.LastUpdateTime = DateTime.Now;
                    }

                    //  新增
                    List<string> insertedIDs = new List<string>();
                    try
                    {
                        insertedIDs = insertRecords.SaveAll();
                    }
                    catch (System.Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.Message);
                        return e.Message;
                    }
                    //  更新
                    List<string> updatedIDs = new List<string>();
                    try
                    {
                        updatedIDs = updateRecords.SaveAll();
                    }
                    catch (System.Exception e)
                    {
                        System.Windows.Forms.MessageBox.Show(e.Message);
                        return e.Message;
                    }

                    //  RaiseEvent
                    if (insertedIDs.Count > 0 || updatedIDs.Count > 0)
                    {
                        //IEnumerable<string> uids = insertedIDs.Union(updatedIDs);
                        UDT.Approach.RaiseAfterUpdateEvent();
                    }

                    ApplicationLog.Log("高雄市國中畢業學生進路調查", "登錄畢業學生進路", "student", "", strLog.ToString());

                    return "儲存成功。";
                }
            }
        }
    }
}
