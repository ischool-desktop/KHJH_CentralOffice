using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Aspose.Words;
using FISCA.UDT;
using KHJHCentralOffice.Properties;

namespace JH_KH_GraduateSurvey.Report
{
    public class CheckReportTemplate
    {
        /// <summary>
        /// 請傳入「填報年度」.
        /// </summary>
        public static MemoryStream Execute(int Year)
        {
            Template Template_102 = new Template_102();
            Template Template_103 = new Template_103();

            //  設定責任鏈之關連
            Template_102.SetSuccessor(Template_103);

            //  開始責任鏈之走訪並回傳結果
            return Template_102.ProcessRequest(Year);
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
                if (Year != 102 && this.successor != null)
                {
                    return this.successor.ProcessRequest(Year);
                }
                else
                {
                    return new MemoryStream(Resources._102學年度國中畢業學生進路調查填報複核表);
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
                if (Year != 103 && this.successor != null)
                {
                    return this.successor.ProcessRequest(Year);
                }
                else
                {
                    return new MemoryStream(Resources._103學年度國中畢業學生進路調查填報複核表);
                }
            }
        }
    }
}
