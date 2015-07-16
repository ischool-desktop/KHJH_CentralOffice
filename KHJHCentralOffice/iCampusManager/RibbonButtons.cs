using System.Collections.Generic;
using System.Windows.Forms;
using FISCA.Presentation;
using FISCA.UDT;

namespace KHJHCentralOffice
{
    class RibbonButtons
    {
        public RibbonButtons()
        {
            Program.MainPanel.RibbonBarItems["管理"]["新增"].Image = Properties.Resources.atom_add_128;
            Program.MainPanel.RibbonBarItems["管理"]["新增"].Size = RibbonBarButton.MenuButtonSize.Large;
            Program.MainPanel.RibbonBarItems["管理"]["新增"].Click += delegate
            {
                DialogResult dr = new AddNewForm().ShowDialog();
                if (dr == DialogResult.OK)
                    Program.RefreshFilteredSource();
            };

            Program.MainPanel.RibbonBarItems["管理"]["刪除"].Image = Properties.Resources.atom_close_128;
            Program.MainPanel.RibbonBarItems["管理"]["刪除"].Size = RibbonBarButton.MenuButtonSize.Large;
            Program.MainPanel.RibbonBarItems["管理"]["刪除"].Click += delegate
            {
                DialogResult dr = MessageBox.Show("刪除選擇的學校？", "Campus", MessageBoxButtons.YesNo);

                if (dr == DialogResult.Yes)
                {
                    AccessHelper ah = new AccessHelper();
                    List<School> schools = ah.Select<School>(Program.MainPanel.SelectedSource);

                    schools.ForEach((x) => x.Deleted = true);
                    schools.SaveAll();
                    Program.RefreshFilteredSource();
                }
            };



            //Program.MainPanel.RibbonBarItems["進階"]["搜尋"].Image = Properties.Resources.lamp_search_128;
            //Program.MainPanel.RibbonBarItems["進階"]["搜尋"].Size = RibbonBarButton.MenuButtonSize.Medium;
            //Program.MainPanel.RibbonBarItems["進階"]["搜尋"].Click += delegate
            //{
            //    new SearchForm().ShowDialog();
            //};

            //Program.MainPanel.RibbonBarItems["進階"]["批次"].Size = RibbonBarButton.MenuButtonSize.Medium;
            //Program.MainPanel.RibbonBarItems["進階"]["批次"]["Desktop 管理"].Click += delegate
            //{
            //    new DesktopModuleManagerForm().ShowDialog();
            //};
            //Program.MainPanel.RibbonBarItems["進階"]["批次"]["UDM 管理"].Click += delegate
            //{
            //    new UDMManagerForm().ShowDialog();
            //};

            //Program.MainPanel.SelectedSourceChanged += delegate
            //{
            //    Program.MainPanel.RibbonBarItems["總務"]["單筆查詢"].Enable = Program.MainPanel.SelectedSource.Count > 0;
            //};
            //Program.MainPanel.RibbonBarItems["總務"]["單筆查詢"].Size = RibbonBarButton.MenuButtonSize.Medium;
            //Program.MainPanel.RibbonBarItems["總務"]["單筆查詢"].Click += delegate
            //{
            //    new VirtualAccountFound.VAFinder().ShowDialog();
            //};
        }
    }
}
