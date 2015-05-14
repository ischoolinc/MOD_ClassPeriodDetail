using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemeritSpecialReport
{
    class StudentObj
    {
        //原始獎懲
        public int MeritA, MeritB, MeritC, DemeritA, DemeritB, DemeritC;
        //功過換算累計獎懲
        public int Result;
        public string ClassName, Name, Id, StudentNo, SeatNo, Grade, DisplayOrder;

        public StudentObj()
        {
            Id = "";
            ClassName = "";
            Name = "";
            StudentNo = "";
            SeatNo = "";
            Grade = "";
            DisplayOrder = "";
            MeritA = 0;
            MeritB = 0;
            MeritC = 0;
            DemeritA = 0;
            DemeritB = 0;
            DemeritC = 0;
            Result = 0;
        }
    }
}
