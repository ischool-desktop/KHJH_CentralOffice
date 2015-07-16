using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Aspose.Words;
using FISCA.UDT;

namespace JH_KH_GraduateSurvey.Accessor
{
    public class ApproachReportTemplate
    {
        /// <summary>
        /// 請傳入「填報年度」.
        /// </summary>
        public static MemoryStream Execute(int Year)
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
            return Template_102.ProcessRequest(ShiftedYear);
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

            public abstract MemoryStream ProcessRequest(int Year);
        }

        /// <summary>
        /// The 'ConcreteHandler' class
        /// </summary>
        private class Template_102 : Template
        {
            public override MemoryStream ProcessRequest(int Year)
            {
                if (Year != 102)
                {
                    return this.successor.ProcessRequest(Year);
                }
                else
                {
                    return new MemoryStream(Properties.Resources.高雄市國中畢業生進路統計報表樣版);
                }
            }
        }

        /// <summary>
        /// The 'ConcreteHandler' class
        /// </summary>
        private class Template_103 : Template
        {
            public override MemoryStream ProcessRequest(int Year)
            {
                if (Year != 103)
                {
                    return this.successor.ProcessRequest(Year);
                }
                else
                {
                    return new MemoryStream(Properties.Resources.高雄市國中畢業生進路統計報表樣版___103);
                }
            }
        }
    }
}
