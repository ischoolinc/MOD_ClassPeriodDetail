using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OverTwiceDemeritA
{
    class Permissions
    {
        public static string 犯過累計滿2大過名單 { get { return "OverTwiceDemeritA.911C3549-2688-4C2D-8898-9813DD18AB72"; } }

        public static bool 犯過累計滿2大過名單權限
        {
            get { return FISCA.Permission.UserAcl.Current[犯過累計滿2大過名單].Executable; }
        }
    }
}
