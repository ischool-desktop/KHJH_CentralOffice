using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using EMBA.DocumentValidator;
using EMBA.Import;
using EMBA.Validator;
using FISCA.Data;
using FISCA.LogAgent;
using FISCA.UDT;

namespace JH_KH_GraduateSurvey.Import
{
    class Approach_Import : ImportWizard
    {
        private StringBuilder strLog = new StringBuilder();
        private ImportOption mOption;
        private AccessHelper Access;
        private QueryHelper Query;
        private List<UDT.Approach> ExistingRecords;
        private Dictionary<string, Dictionary<string, string>> dicStudents;
        private int SchemaYear;
        private int SchoolYear;
        private List<string> SurveyFields;

        public Approach_Import(int SchoolYear)
        {
            this.IsSplit = false;
            this.ShowAdvancedForm = false;
            this.ValidateRuleFormater = XDocument.Parse(Properties.Resources.format);

            this.SchoolYear = SchoolYear;
            this.Access = new AccessHelper();
            this.Query = new QueryHelper();
            this.ExistingRecords = new List<UDT.Approach>();
            this.dicStudents = new Dictionary<string, Dictionary<string, string>>();
            this.SurveyFields = new List<string>();

            this.CustomValidate = (Rows, Messages) =>
            {
                CustomValidator(Rows, Messages);
            };
        }

        public void CustomValidator(List<IRowStream> Rows, RowMessages Messages)
        {
            DataTable dataTables = this.Query.Select("select id_number, id, name from student");
            foreach (DataRow row in dataTables.Rows)
            {
                if (string.IsNullOrWhiteSpace(row["id_number"] + ""))
                    continue;

                if (!this.dicStudents.ContainsKey(row["id_number"] + ""))
                    this.dicStudents.Add(row["id_number"] + "", new Dictionary<string, string>() {{ row["id"] + "", row["name"] + "" }});
            }
            Dictionary<int, Dictionary<string, string>> Data = new Dictionary<int, Dictionary<string, string>>();
            Rows.ForEach((x) =>
            {
                Data.Add(x.Position, new Dictionary<string, string>());
                foreach(string Field in this.SurveyFields)
                {
                    Data[x.Position].Add(Field, x.GetValue(Field));
                }
            });
            Rows.ForEach((x) =>
            {
                string id_number = x.GetValue("身分證號").Trim();
                string name = x.GetValue("姓名").Trim();
                #region 檢查身分證號
                //  「身分證號」必須存在於系統
                if (!this.dicStudents.ContainsKey(id_number))
                {
                    Messages[x.Position].MessageItems.Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "身分證號不存在。"));
                }
                //  「身分證號」必須與「姓名」一致
                else
                {
                    if (this.dicStudents[id_number].ElementAt(0).Value.Trim().ToLower() != name.ToLower())
                        Messages[x.Position].MessageItems.Add(new MessageItem(EMBA.Validator.ErrorType.Error, EMBA.Validator.ValidatorType.Row, "使用身分證號驗證學生姓名錯誤。"));
                }
                #endregion
            });
            Dictionary<int, List<MessageItem>> dicMessages = Accessor.ApproachValidate.Execute(this.SchoolYear, Data);
            foreach(int key in dicMessages.Keys)
            {
                Messages[key].MessageItems.AddRange(dicMessages[key]);
            }
        }

        public override ImportAction GetSupportActions()
        {
            return ImportAction.InsertOrUpdate;
        }

        public override XDocument GetValidateRule()
        {            
            try
            {
                XDocument document = XDocument.Parse(Properties.Resources.Approach_Import, LoadOptions.None);
                XElement element = document.Descendants("Schema").OrderByDescending(x=>int.Parse(x.Attribute("SchoolYear").Value)).First();
                this.SchemaYear = document.Descendants("Schema").Max(x => int.Parse(x.Attribute("SchoolYear").Value));
                document.Descendants("FieldList").First().Descendants("Field").LastOrDefault().AddAfterSelf(element.Element("FieldList").Elements());
                document.Descendants("ValidatorList").First().Add(element.Element("ValidatorList").Elements());
                this.SurveyFields.AddRange(element.Element("FieldList").Elements().Select(x => x.Attribute("Name").Value));
                return document;
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }

        public override string Import(List<IRowStream> Rows)
        {
            Dictionary<string, Dictionary<string, string>> Data = new Dictionary<string, Dictionary<string, string>>();
            foreach (IRowStream row in Rows)
            {
                string id_number = row.GetValue("身分證號").Trim();
                string student_id = this.dicStudents[id_number].ElementAt(0).Key;

                Data.Add(student_id, new Dictionary<string, string>());
                foreach(string Field in this.SurveyFields)
                {
                    Data[student_id].Add(Field, row.GetValue(Field).Trim());
                }
            }
            string Message = Accessor.ApproachSave.Execute(this.SchoolYear, Data);
            return Message;
        }

        public override void Prepare(ImportOption Option)
        {
            mOption = Option;
        }
    }
}
