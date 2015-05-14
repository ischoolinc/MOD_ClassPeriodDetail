using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OverTwiceDemeritA
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["學務作業", "資料統計"];
            item1["報表"].Image = Properties.Resources.Report;
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            item1["報表"]["大過累計達標準名單(客製)"].Enable = Permissions.犯過累計滿2大過名單權限;
            item1["報表"]["大過累計達標準名單(客製)"].Click += delegate
            {
                Printer printer = new Printer();
                printer.ShowDialog();
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["學務作業"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.犯過累計滿2大過名單, "大過累計達標準名單(客製)"));
        }
    }
}
