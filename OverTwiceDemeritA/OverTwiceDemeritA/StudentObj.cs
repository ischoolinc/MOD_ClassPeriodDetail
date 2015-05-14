using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OverTwiceDemeritA
{
    class StudentObj
    {
        //原始獎懲
        public int MeritA, MeritB, MeritC, DemeritA, DemeritB, DemeritC;
        //功過換算累計獎懲
        public int MA, MB, MC, DA, DB, DC;
        public string ClassName, Name, Id, StudentNo, SeatNo, Status, Grade, DisplayOrder;
        public bool 留察;
        public StudentObj(string id)
        {
            Id = id;
            StudentNo="";
            SeatNo="";
            ClassName = "";
            Name = "";

            MeritA=0; 
            MeritB=0; 
            MeritC=0;
            DemeritA = 0;
            DemeritB = 0;
            DemeritC = 0;
            MA = 0;
            MB = 0;
            MC = 0;
            DA = 0;
            DB = 0;
            DC = 0;
            
            留察 = false;
        }
    }
}
