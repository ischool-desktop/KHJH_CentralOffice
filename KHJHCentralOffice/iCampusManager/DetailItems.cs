using DesktopLib;

namespace KHJHCentralOffice
{
    internal class DetailItems
    {
        public DetailItems()
        {
            Program.MainPanel.RegisterDetailContent<BasicInfoItem>();
            //Program.MainPanel.RegisterDetailContent<GraduateSurveyVagrant>();
            Program.MainPanel.RegisterDetailContent<GraduateSurveyApproach>();

            //Program.MainPanel.RegisterDetailContent<NetworkItem>();
            //Program.MainPanel.RegisterDetailContent<UDMItem>();
            //Program.MainPanel.RegisterDetailContent<UDTItem>();
            //Program.MainPanel.RegisterDetailContent<UDSItem>();
            //Program.MainPanel.RegisterDetailContent<ModuleItem>();
            //Program.MainPanel.RegisterDetailContent<WebGadgetItem>();
            //Program.MainPanel.RegisterDetailContent<StudentExchangeReset>();
        }
    }
}