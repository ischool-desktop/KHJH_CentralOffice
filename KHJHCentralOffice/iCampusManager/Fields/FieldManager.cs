using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KHJHCentralOffice
{
    internal class FieldManager
    {
        internal static TitleField TitleField { get; private set; }

        internal static DSNSField DSNSField { get; private set; }

        internal static GroupField GroupField { get; private set; }

        public FieldManager()
        {

            TitleField = new KHJHCentralOffice.TitleField();
            GroupField = new KHJHCentralOffice.GroupField();
            DSNSField = new KHJHCentralOffice.DSNSField();

            TitleField.Register(Program.MainPanel);
            //DSNSField.Register(Program.MainPanel);
            GroupField.Register(Program.MainPanel);
        }
    }
}