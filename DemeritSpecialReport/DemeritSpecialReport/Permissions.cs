using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemeritSpecialReport
{
    class Permissions
    {
        public static string 懲戒特殊表現 { get { return "DemeritSpecialReport.185751EC-5472-425E-8542-EA276F3B3713"; } }

        public static bool 懲戒特殊表現權限
        {
            get { return FISCA.Permission.UserAcl.Current[懲戒特殊表現].Executable; }
        }
    }
}
