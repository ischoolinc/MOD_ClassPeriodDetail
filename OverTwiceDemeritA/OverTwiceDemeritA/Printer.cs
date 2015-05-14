using Aspose.Cells;
using FISCA.Data;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using K12.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace OverTwiceDemeritA
{
    public partial class Printer : BaseForm
    {
        int _SchoolYear,_StandardValue;
        Dictionary<string, ClassRecord> _ClassCatch;
        string _SchoolName;
        string _ReportName;
        BackgroundWorker _BW;
        QueryHelper _Q;
        XmlDocument _XD;
        public Printer()
        {
            InitializeComponent();
        }

        private void Printer_Load(object sender, EventArgs e)
        {
            _SchoolName = K12.Data.School.ChineseName;
            _ReportName = "犯過累計滿2次大過學生名單";
            _XD = new XmlDocument();
            _Q = new QueryHelper();
            _BW = new BackgroundWorker();
            _BW.WorkerReportsProgress = true;
            _BW.DoWork += new DoWorkEventHandler(DataBuilding);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ReportBuilding);
            _BW.ProgressChanged += new ProgressChangedEventHandler(BW_Progress);

            string schoolYear = K12.Data.School.DefaultSchoolYear;
            int year = 0;
            bool isNum = int.TryParse(schoolYear, out year);

            if (isNum)
            {
                for (int i = -2; i < 3; i++)
                {
                    cboYear.Items.Add(year + i);
                }
            }

            cboYear.Text = schoolYear;
        }

        private void BW_Progress(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage(_ReportName + "產生中", e.ProgressPercentage);
        }

        private void ReportBuilding(object sender, RunWorkerCompletedEventArgs e)
        {
            MotherForm.SetStatusBarMessage(_ReportName + " 產生完成");

            EnableForm(true);
            Workbook wb = (Workbook)e.Result;
            SaveFileDialog sd = new SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = _ReportName + ".xls";
            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    wb.Save(sd.FileName, Aspose.Cells.SaveFormat.Excel97To2003);
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    MessageBox.Show("檔案儲存失敗");
                }
            }
        }

        private void DataBuilding(object sender, DoWorkEventArgs e)
        {
            //取得班級物件對照
            _BW.ReportProgress(0);
            List<ClassRecord> classes = K12.Data.Class.SelectAll();
            _ClassCatch = new Dictionary<string, ClassRecord>();
            foreach (ClassRecord record in classes)
            {
                if (!_ClassCatch.ContainsKey(record.ID))
                    _ClassCatch.Add(record.ID, record);
            }

            //查詢指定學年度的獎懲紀錄
            _BW.ReportProgress(10);
            Dictionary<string,StudentObj> StudentDic = new Dictionary<string,StudentObj>();
            DataTable dt = _Q.Select("SELECT school_year,ref_student_id,merit_flag,xpath_string(detail,'/Discipline/Merit/@A') as merita,xpath_string(detail,'/Discipline/Merit/@B') as meritb,xpath_string(detail,'/Discipline/Merit/@C') as meritc,xpath_string(detail,'/Discipline/Demerit/@A') as demerita,xpath_string(detail,'/Discipline/Demerit/@B') as demeritb,xpath_string(detail,'/Discipline/Demerit/@C') as demeritc,xpath_string(detail,'/Discipline/Demerit/@Cleared') as cleared FROM discipline WHERE school_year=" + _SchoolYear);
            foreach (DataRow drow in dt.Rows)
            {
                string sid = drow["ref_student_id"].ToString();
                if (!StudentDic.ContainsKey(sid))
                {
                    StudentDic.Add(sid, new StudentObj(sid));
                }

                int x = 0;
                int merita = int.TryParse(drow["merita"].ToString(),out x) ? x : 0;
                int meritb = int.TryParse(drow["meritb"].ToString(), out x) ? x : 0;
                int meritc = int.TryParse(drow["meritc"].ToString(), out x) ? x : 0;
                int demerita = int.TryParse(drow["demerita"].ToString(), out x) ? x : 0;
                int demeritb = int.TryParse(drow["demeritb"].ToString(), out x) ? x : 0;
                int demeritc = int.TryParse(drow["demeritc"].ToString(), out x) ? x : 0;
                string cleared = drow["cleared"].ToString();
                string flag = drow["merit_flag"].ToString();

                StudentDic[sid].MeritA += merita;
                StudentDic[sid].MeritB += meritb;
                StudentDic[sid].MeritC += meritc;

                if (cleared != "是")
                {
                    StudentDic[sid].DemeritA += demerita;
                    StudentDic[sid].DemeritB += demeritb;
                    StudentDic[sid].DemeritC += demeritc;
                }

                if (flag == "2")
                    StudentDic[sid].留察 = true;
            }

            //功過換算
            _BW.ReportProgress(20);
            MeritDemeritReduceRecord mdrr = MeritDemeritReduce.Select();
            int MAB = mdrr.MeritAToMeritB.HasValue ? mdrr.MeritAToMeritB.Value : 0;
            int MBC = mdrr.MeritBToMeritC.HasValue ? mdrr.MeritBToMeritC.Value : 0;
            int DAB = mdrr.DemeritAToDemeritB.HasValue ? mdrr.DemeritAToDemeritB.Value : 0;
            int DBC = mdrr.DemeritBToDemeritC.HasValue ? mdrr.DemeritBToDemeritC.Value : 0;

            foreach (StudentObj obj in StudentDic.Values)
            {
                int merit = ((obj.MeritA * MAB) + obj.MeritB) * MBC + obj.MeritC;
                int demerit = ((obj.DemeritA * DAB) + obj.DemeritB) * DBC + obj.DemeritC;

                int total = merit - demerit;

                if (total > 0)
                {
                    obj.MC = total % MBC;
                    obj.MB = (total / MBC) % MAB;
                    obj.MA = (total / MBC) / MAB;
                }
                else if (total < 0)
                {
                    total *= -1;
                    obj.DC = total % DBC;
                    obj.DB = (total / DBC) % DAB;
                    obj.DA = (total / DBC) / DAB;
                }
            }

            //取得累積大過滿2次以上的學生
            _BW.ReportProgress(30);
            Dictionary<string, StudentObj> NewStudentDic = new Dictionary<string, StudentObj>();
            foreach (StudentObj obj in StudentDic.Values)
            {
                if (obj.DA >= _StandardValue)
                    NewStudentDic.Add(obj.Id, obj);
            }

            //取得學生資料
            _BW.ReportProgress(40);
            List<string> ids = NewStudentDic.Keys.ToList();
            List<StudentRecord> studentRecords = K12.Data.Student.SelectByIDs(ids);
            foreach (StudentRecord record in studentRecords)
            {
                if (NewStudentDic.ContainsKey(record.ID))
                {
                    NewStudentDic[record.ID].StudentNo = record.StudentNumber;
                    NewStudentDic[record.ID].Name = record.Name;
                    NewStudentDic[record.ID].SeatNo = record.SeatNo.HasValue ? record.SeatNo.Value.ToString() : "";

                    string classId = record.RefClassID;
                    NewStudentDic[record.ID].ClassName = _ClassCatch.ContainsKey(classId) ? _ClassCatch[classId].Name : "";
                    NewStudentDic[record.ID].Grade = _ClassCatch.ContainsKey(classId) ? _ClassCatch[classId].GradeYear.ToString() : "";
                    NewStudentDic[record.ID].DisplayOrder = _ClassCatch.ContainsKey(classId) ? _ClassCatch[classId].DisplayOrder : "";

                    NewStudentDic[record.ID].Status = record.Status.ToString();
                }
            }

            //列印
            _BW.ReportProgress(50);
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.Template));
            Worksheet ws = wb.Worksheets[0];
            Cells cs = ws.Cells;

            //需動態新增的框線
            _BW.ReportProgress(60);
            Range eachLine = cs.CreateRange(3, 1, false);

            //Report Title
            _BW.ReportProgress(70);
            cs[0, 0].PutValue(_SchoolName + " " + _ReportName);

            //資料排序
            _BW.ReportProgress(80);
            List<StudentObj> DataList = NewStudentDic.Values.ToList();
            DataList.Sort(DataListSort);

            int row = 2;
            foreach (StudentObj obj in DataList)
            {
                if (obj.Status == "一般" || obj.Status == "延修")
                {
                    cs.CreateRange(row + 1, 1, false).CopyStyle(eachLine);
                    cs[row, 0].PutValue(obj.StudentNo);
                    cs[row, 1].PutValue(obj.ClassName);
                    cs[row, 2].PutValue(obj.SeatNo);
                    cs[row, 3].PutValue(obj.Name);
                    cs[row, 4].PutValue(obj.留察 ? "Y" : "");
                    cs[row, 5].PutValue(obj.MA == 0 ? "" : obj.MA.ToString());
                    cs[row, 6].PutValue(obj.MB == 0 ? "" : obj.MB.ToString());
                    cs[row, 7].PutValue(obj.MC == 0 ? "" : obj.MC.ToString());
                    cs[row, 8].PutValue(obj.DA == 0 ? "" : obj.DA.ToString());
                    cs[row, 9].PutValue(obj.DB == 0 ? "" : obj.DB.ToString());
                    cs[row, 10].PutValue(obj.DC == 0 ? "" : obj.DC.ToString());
                    row++;
                }
            }

            e.Result = wb;
            _BW.ReportProgress(100);
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (_BW.IsBusy)
            {
                MsgBox.Show("系統忙碌請稍後再試...");
            }
            else
            {
                EnableForm(false);
                _SchoolYear = int.Parse(cboYear.Text);
                _StandardValue = (int)numericUpDown1.Value;
                _ReportName = "犯過累計滿" + _StandardValue + "次大過學生名單";
                _BW.RunWorkerAsync();
            }
        }

        private void EnableForm(bool p)
        {
            this.cboYear.Enabled = p;
            this.buttonX1.Enabled = p;
            this.buttonX2.Enabled = p;
        }

        private int DataListSort(StudentObj x, StudentObj y)
        {
            string xx = x.Grade.PadLeft(2, '0');
            xx += x.DisplayOrder.PadLeft(5, '0');
            xx += x.ClassName.PadLeft(20, '0');
            xx += x.SeatNo.PadLeft(3, '0');

            string yy = y.Grade.PadLeft(2, '0');
            yy += y.DisplayOrder.PadLeft(5, '0');
            yy += y.ClassName.PadLeft(20, '0');
            yy += y.SeatNo.PadLeft(3, '0');

            return xx.CompareTo(yy);
        }
    }
}
