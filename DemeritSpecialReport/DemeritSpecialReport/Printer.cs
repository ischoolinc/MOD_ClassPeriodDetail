using Aspose.Cells;
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

namespace DemeritSpecialReport
{
    public partial class Printer : BaseForm
    {
        List<string> _Classes;
        string _SchoolName, _ReportName;
        int _SchoolYear, _Semester, _StandardValue;
        BackgroundWorker _BW;
        public Printer(List<string> source)
        {
            InitializeComponent();
            _Classes = source;
        }

        private void Printer_Load(object sender, EventArgs e)
        {
            _SchoolName = K12.Data.School.ChineseName;
            _ReportName = "懲戒特殊表現";
            _BW = new BackgroundWorker();
            _BW.WorkerReportsProgress = true;
            _BW.DoWork += new DoWorkEventHandler(DataBuilding);
            _BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(ReportBuilding);
            _BW.ProgressChanged += new ProgressChangedEventHandler(BW_Progress);

            string schoolYear = K12.Data.School.DefaultSchoolYear;
            string semester = K12.Data.School.DefaultSemester;
            int year = 0;
            bool isNum = int.TryParse(schoolYear, out year);

            if (isNum)
            {
                for (int i = -2; i < 3; i++)
                {
                    cboSchoolYear.Items.Add(year + i);
                }
            }
            else
            {
                MessageBox.Show("系統預設學年度不正確,請確認值為數字");
                this.Close();
            }

            cboSemester.Items.Add("1");
            cboSemester.Items.Add("2");

            cboSchoolYear.Text = schoolYear;
            cboSemester.Text = semester;
        }

        private void DataBuilding(object sender, DoWorkEventArgs e)
        {
            _BW.ReportProgress(0);
            //班級資料
            List<ClassRecord> classRecords = K12.Data.Class.SelectByIDs(_Classes);

            Dictionary<string, ClassRecord> classCatch = new Dictionary<string, ClassRecord>();
            Dictionary<string, StudentObj> Students = new Dictionary<string, StudentObj>();

            //建立班級快取並展開學生資料
            _BW.ReportProgress(10);
            foreach (ClassRecord record in classRecords)
            {
                string id = record.ID;
                if (!classCatch.ContainsKey(id))
                    classCatch.Add(id, record);

                foreach (StudentRecord student in record.Students)
                {
                    if (student.Status == StudentRecord.StudentStatus.一般 || student.Status == StudentRecord.StudentStatus.延修)
                    {
                        string sid = student.ID;
                        if (!Students.ContainsKey(sid))
                            Students.Add(sid, new StudentObj());

                        Students[sid].Id = student.ID;
                        Students[sid].Name = student.Name;
                        Students[sid].SeatNo = student.SeatNo.HasValue ? student.SeatNo.ToString() : "";
                        Students[sid].StudentNo = student.StudentNumber;

                        if (classCatch.ContainsKey(student.RefClassID))
                        {
                            Students[sid].ClassName = classCatch[student.RefClassID].Name;
                            Students[sid].Grade = classCatch[student.RefClassID].GradeYear.HasValue ? classCatch[student.RefClassID].GradeYear.Value.ToString() : "";
                            Students[sid].DisplayOrder = classCatch[student.RefClassID].DisplayOrder;
                        }
                    }
                }
            }

            //獎懲紀錄
            _BW.ReportProgress(20);
            List<DisciplineRecord> disciplineRecords = Discipline.SelectByStudentIDs(Students.Keys);
            foreach (DisciplineRecord dr in disciplineRecords)
            {
                if (dr.SchoolYear != _SchoolYear) continue;

                if (_Semester == 1)
                    if (dr.Semester != 1) continue;

                string id = dr.RefStudentID;
                Students[id].MeritA += dr.MeritA.HasValue ? dr.MeritA.Value : 0;
                Students[id].MeritB += dr.MeritB.HasValue ? dr.MeritB.Value : 0;
                Students[id].MeritC += dr.MeritC.HasValue ? dr.MeritC.Value : 0;

                if (dr.Cleared != "是")
                {
                    Students[id].DemeritA += dr.DemeritA.HasValue ? dr.DemeritA.Value : 0;
                    Students[id].DemeritB += dr.DemeritB.HasValue ? dr.DemeritB.Value : 0;
                    Students[id].DemeritC += dr.DemeritC.HasValue ? dr.DemeritC.Value : 0;
                }
            }

            //功過換算
            _BW.ReportProgress(30);
            MeritDemeritReduceRecord mdrr = MeritDemeritReduce.Select();
            int MAB = mdrr.MeritAToMeritB.HasValue ? mdrr.MeritAToMeritB.Value : 0;
            int MBC = mdrr.MeritBToMeritC.HasValue ? mdrr.MeritBToMeritC.Value : 0;
            int DAB = mdrr.DemeritAToDemeritB.HasValue ? mdrr.DemeritAToDemeritB.Value : 0;
            int DBC = mdrr.DemeritBToDemeritC.HasValue ? mdrr.DemeritBToDemeritC.Value : 0;

            foreach (StudentObj obj in Students.Values)
            {
                int merit = ((obj.MeritA * MAB) + obj.MeritB) * MBC + obj.MeritC;
                int demerit = ((obj.DemeritA * DAB) + obj.DemeritB) * DBC + obj.DemeritC;

                int total = merit - demerit;

                if (total < 0)
                {
                    total *= -1;
                    obj.Result = total;
                }
            }

            //資料排序
            _BW.ReportProgress(40);
            List<StudentObj> PrintList = Students.Values.ToList();
            PrintList.Sort(SortObj);

            //預備填入資料
            _BW.ReportProgress(50);
            Workbook wb = new Workbook(new MemoryStream(Properties.Resources.Template));

            Cells cs = wb.Worksheets[0].Cells;

            cs[0, 0].PutValue(_SchoolName + " (" + _SchoolYear + "/" + _Semester + ") " + _ReportName);

            //開始填入
            decimal per = (decimal)50 / PrintList.Count;
            int count = 0;
            int row = 2;
            foreach (StudentObj obj in PrintList)
            {
                if (obj.Result >= _StandardValue)
                {
                    cs[row, 0].PutValue(obj.ClassName);
                    cs[row, 1].PutValue(obj.SeatNo);
                    cs[row, 2].PutValue(obj.Name);
                    cs[row, 3].PutValue(obj.StudentNo);
                    cs[row, 4].PutValue(obj.DemeritA);
                    cs[row, 5].PutValue(obj.DemeritB);
                    cs[row, 6].PutValue(obj.DemeritC);
                    cs[row, 7].PutValue(obj.MeritA);
                    cs[row, 8].PutValue(obj.MeritB);
                    cs[row, 9].PutValue(obj.MeritC);
                    cs[row, 10].PutValue(obj.Result);
                    row++;
                }

                _BW.ReportProgress((int)per*count + 50);
                count++;
            }

            _BW.ReportProgress(100);
            e.Result = wb;
        }

        private int SortObj(StudentObj x, StudentObj y)
        {
            string xx = x.Grade.PadLeft(2, '0');
            xx += x.DisplayOrder != "" ? x.DisplayOrder.PadLeft(3,'0') : "999";
            xx += x.ClassName.PadLeft(20, '0');
            xx += x.SeatNo.PadLeft(3, '0');

            string yy = y.Grade.PadLeft(2, '0');
            yy += y.DisplayOrder != "" ? y.DisplayOrder.PadLeft(3, '0') : "999";
            yy += y.ClassName.PadLeft(20, '0');
            yy += y.SeatNo.PadLeft(3, '0');

            return xx.CompareTo(yy);
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

        private void BW_Progress(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage(_ReportName + " 產生中", e.ProgressPercentage);
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
                _SchoolYear = int.Parse(cboSchoolYear.Text);
                _Semester = int.Parse(cboSemester.Text);
                _StandardValue = (int)numericUpDown1.Value;
                _BW.RunWorkerAsync();
            }
        }

        private void EnableForm(bool p)
        {
            this.cboSchoolYear.Enabled = p;
            this.cboSemester.Enabled = p;
            this.numericUpDown1.Enabled = p;
            this.buttonX1.Enabled = p;
            this.buttonX2.Enabled = p;
        }


    }
}
