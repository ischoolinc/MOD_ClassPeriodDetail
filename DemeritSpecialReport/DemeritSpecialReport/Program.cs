using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemeritSpecialReport
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["班級", "資料統計"];
            item1["報表"].Image = Properties.Resources.Report;
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                item1["報表"]["學務相關報表"]["懲戒特殊表現(功過相抵)"].Enable = Permissions.懲戒特殊表現權限 && (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0);
            };

            item1["報表"]["學務相關報表"]["懲戒特殊表現(功過相抵)"].Click += delegate
            {
                Printer printer = new Printer(K12.Presentation.NLDPanels.Class.SelectedSource);
                printer.ShowDialog();
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["班級"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.懲戒特殊表現, "懲戒特殊表現(功過相抵)"));
        }
    }
}
