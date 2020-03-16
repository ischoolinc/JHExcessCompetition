using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using System.Data;
using System.Xml.Linq;
using JHSchool.Data;

namespace ChiaYiExcessCompetition.DAO
{
    public class QueryData
    {
        /// <summary>
        /// 取得三年級一般狀態學生資料
        /// </summary>
        /// <returns></returns>
        public static List<StudentInfo> GetStudentInfoList3()
        {
            List<StudentInfo> value = new List<StudentInfo>();
            QueryHelper qh = new QueryHelper();
            string qry = @"
SELECT 
    student.id AS student_id
    ,class.id AS class_id
    ,student_number
    ,class.class_name
    ,seat_no
    ,student.name AS student_name
    ,id_number
    ,birthdate
    ,CASE gender WHEN '0' THEN '2' WHEN '1' THEN '1' ELSE '' END AS gender
    ,sems_history
FROM 
student INNER JOIN class 
ON student.ref_class_id = class.id 
WHERE student.status = 1 AND class.grade_year IN(3,9) 
ORDER BY class.grade_year,class.display_order,class.class_name,seat_no
";
            DataTable dt = qh.Select(qry);
            if (dt != null)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    StudentInfo si = new StudentInfo();
                    si.StudentID = dr["student_id"].ToString();
                    si.ClassID = dr["class_id"].ToString();
                    si.StudentNumber = dr["student_number"].ToString();

                    if (K12.Data.School.Code.Length >= 6)
                    {
                        si.SchoolCode = K12.Data.School.Code.Substring(0, 6);
                    }
                    else
                        si.SchoolCode = K12.Data.School.Code;

                    string className = dr["class_name"].ToString();

                    // 班名取2位
                    if (className.Length >= 2)
                        si.ClassName = className.Substring(className.Length - 2, 2);
                    else
                        si.ClassName = className;

                    // 座號補0
                    int no;
                    if (int.TryParse(dr["seat_no"].ToString(), out no))
                    {
                        if (no < 10)
                            si.SeatNo = "0" + no;
                        else
                            si.SeatNo = no + "";
                    }
                    else
                    {
                        si.SeatNo = "";
                    }

                    si.StudentName = dr["student_name"].ToString();


                    si.IDNumber = dr["id_number"].ToString().Trim();

                    // 檢查是否台灣身分證
                    string ii = si.IDNumber.Substring(1, 1);
                    if (ii == "1" || ii == "2")
                    {
                        si.isTaiwanID = true;
                    }
                    else
                    {
                        si.isTaiwanID = false;
                    }

                    si.GenderCode = dr["gender"].ToString();
                    DateTime dt1;
                    if (DateTime.TryParse(dr["birthdate"].ToString(), out dt1))
                    {
                        si.BirthYear = (dt1.Year - 1911) + "";
                        si.BirthMonth = dt1.Month + "";
                        si.BirthDay = dt1.Day + "";
                    }

                    // 處理學期歷程
                    if (dr["sems_history"] != null)
                    {
                        try
                        {
                            XElement elms = XElement.Parse("<root>" + dr["sems_history"].ToString() + "</root>");
                            foreach (XElement elm in elms.Elements("History"))
                            {
                                SemsHistoryInfo shi = new SemsHistoryInfo();
                                if (elm.Attribute("SchoolYear") != null)
                                    shi.SchoolYear = elm.Attribute("SchoolYear").Value;

                                if (elm.Attribute("Semester") != null)
                                    shi.Semester = elm.Attribute("Semester").Value;

                                if (elm.Attribute("GradeYear") != null)
                                    shi.GradeYear = elm.Attribute("GradeYear").Value;

                                si.SemsHistoryInfoList.Add(shi);
                            }

                        }
                        catch (Exception ex) { }

                    }

                    value.Add(si);
                }
            }
            return value;
        }

        /// <summary>
        /// 取得學生 return idprefix:name
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, List<string>> GetStudentTagName(List<string> studIDList)
        {

            Dictionary<string, List<string>> value = new Dictionary<string, List<string>>();

            if (studIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string qry = @"SELECT 
ref_student_id AS student_id
,(CASE WHEN prefix is null THEN name ELSE prefix||':'||name END) AS tag_name 
FROM 
tag INNER JOIN tag_student ON tag.id = tag_student.ref_tag_id
WHERE ref_student_id IN(" + string.Join(",", studIDList.ToArray()) + @") AND category =  'Student' ORDER BY ref_student_id,prefix,name
";

                DataTable dt = qh.Select(qry);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string id = dr["student_id"].ToString();
                        string tagName = dr["tag_name"].ToString();

                        if (!value.ContainsKey(id))
                            value.Add(id, new List<string>());

                        value[id].Add(tagName);
                    }
                }
            }
            return value;
        }

        /// <summary>
        /// 取得中低收入戶並填入
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="StudentInfoList"></param>
        /// <returns></returns>

        public static List<StudentInfo> FillIncomeType(List<string> StudentIDList, List<StudentInfo> StudentInfoList)
        {
            try
            {
                // 取得中低收入資料
                QueryHelper qh = new QueryHelper();

                Dictionary<string, DataRow> compDict = new Dictionary<string, DataRow>();
                string qry = @"
SELECT ref_student_id AS student_id
,income_type 
FROM student_info_ext WHERE ref_student_id IN(" + string.Join(",", StudentIDList.ToArray()) + @")

";
                DataTable dt = qh.Select(qry);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["student_id"].ToString();
                        if (!compDict.ContainsKey(sid))
                            compDict.Add(sid, dr);
                    }
                }

                // 填入資料
                foreach (StudentInfo si in StudentInfoList)
                {
                    if (compDict.ContainsKey(si.StudentID))
                    {
                        if (compDict[si.StudentID]["income_type"] != null)
                        {
                            if (compDict[si.StudentID]["income_type"].ToString() == "低")
                                si.incomeType1 = true;

                            if (compDict[si.StudentID]["income_type"].ToString() == "中低")
                                si.incomeType2 = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // throw ex;
            }

            return StudentInfoList;
        }

        /// <summary>
        /// 取得學生體適能資料並填入
        /// </summary>
        /// <param name="StudentInfoList"></param>
        /// <param name="endDate"></param>
        public static List<StudentInfo> FillStudentFitness(List<string> StudentIDList, List<StudentInfo> StudentInfoList, DateTime endDate)
        {
            try
            {

                List<string> addStringList = new List<string>();
                addStringList.Add("金牌");
                addStringList.Add("銀牌");
                addStringList.Add("銅牌");

                // 嘉義版先不卡日期
                // 取得特定日期前體適能資料
                QueryHelper qh = new QueryHelper();
                endDate = endDate.AddDays(1);
                string strEndDate = endDate.Year + "-" + endDate.Month + "-" + endDate.Day;
                Dictionary<string, List<DataRow>> finDict = new Dictionary<string, List<DataRow>>();
                //                string qry = @"
                //SELECT 
                //ref_student_id
                //,test_date
                //,sit_and_reach_degree
                //,standing_long_jump_degree
                //,sit_up_degree
                //,cardiorespiratory_degree
                //FROM $ischool_student_fitness WHERE test_date <'" + strEndDate + @"' AND ref_student_id IN('" + string.Join("','", StudentIDList.ToArray()) + @"')
                //";
                string qry = @"
SELECT 
ref_student_id
,test_date
,sit_and_reach_degree
,standing_long_jump_degree
,sit_up_degree
,cardiorespiratory_degree
FROM $ischool_student_fitness WHERE ref_student_id IN('" + string.Join("','", StudentIDList.ToArray()) + @"')
";

                DataTable dt = qh.Select(qry);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["ref_student_id"].ToString();
                        if (!finDict.ContainsKey(sid))
                            finDict.Add(sid, new List<DataRow>());

                        finDict[sid].Add(dr);
                    }
                }

                // 填入體適能資料
                foreach (StudentInfo si in StudentInfoList)
                {
                    if (finDict.ContainsKey(si.StudentID))
                    {
                        si.isAddFitnessScore = false;

                        foreach (DataRow dr in finDict[si.StudentID])
                        {
                            int addCount = 0;

                            // sit_and_reach_degree
                            if (dr["sit_and_reach_degree"] != null)
                            {
                                string ss = dr["sit_and_reach_degree"].ToString().Trim();
                                if (ss == "" || ss == "未檢測")
                                {

                                }
                                else
                                {
                                    si.sit_and_reach_degreeList.Add(ss);
                                    if (addStringList.Contains(ss))
                                        addCount++;
                                }
                            }

                            // standing_long_jump_degree
                            if (dr["standing_long_jump_degree"] != null)
                            {
                                string ss = dr["standing_long_jump_degree"].ToString().Trim();
                                if (ss == "" || ss == "未檢測")
                                {

                                }
                                else
                                {
                                    si.standing_long_jump_degreeList.Add(ss);
                                    if (addStringList.Contains(ss))
                                        addCount++;
                                }
                            }

                            // sit_up_degree
                            if (dr["sit_up_degree"] != null)
                            {
                                string ss = dr["sit_up_degree"].ToString().Trim();
                                if (ss == "" || ss == "未檢測")
                                {

                                }
                                else
                                {
                                    si.sit_up_degreeList.Add(ss);
                                    if (addStringList.Contains(ss))
                                        addCount++;
                                }
                            }

                            // cardiorespiratory_degree
                            if (dr["cardiorespiratory_degree"] != null)
                            {
                                string ss = dr["cardiorespiratory_degree"].ToString().Trim();
                                if (ss == "" || ss == "未檢測")
                                {

                                }
                                else
                                {
                                    si.cardiorespiratory_degreeList.Add(ss);
                                    if (addStringList.Contains(ss))
                                        addCount++;
                                }
                            }

                            // 4 次都符合銅牌以上，符合加分
                            if (addCount == 4)
                                si.isAddFitnessScore = true;

                        }

                    }
                }

            }
            catch (Exception ex)
            {
                //  throw ex;
            }

            return StudentInfoList;
        }

        /// <summary>
        /// 取得學生所有可選類別 return prefix:name,id
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetStudentAllTag()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            string qry = @"SELECT 
id
,(CASE WHEN prefix is null THEN name ELSE prefix||':'||name END) AS tag_name 
FROM 
tag 
WHERE category =  'Student' ORDER BY prefix,name";

            DataTable dt = qh.Select(qry);
            if (dt != null)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr["id"].ToString();
                    string tagName = dr["tag_name"].ToString();

                    if (!value.ContainsKey(tagName))
                        value.Add(tagName, id);
                }
            }
            return value;
        }


        /// <summary>
        /// 取得服務學習時數並填入
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="StudentInfoList"></param>
        /// <returns></returns>
        public static List<StudentInfo> FillServiceLearn(List<string> StudentIDList, List<StudentInfo> StudentInfoList, DateTime endDate)
        {
            Dictionary<string, decimal> srDict = new Dictionary<string, decimal>();
            endDate = endDate.AddDays(1);
            string strEndDate = endDate.Year + "-" + endDate.Month + "-" + endDate.Day;

            try
            {
                QueryHelper qh = new QueryHelper();
                string qry = @"
SELECT ref_student_id AS student_id,sum(hours) AS hour_sum FROM $k12.service.learning.record WHERE occur_date < '" + strEndDate + @"' AND ref_student_id IN('" + string.Join("','", StudentIDList.ToArray()) + @"') 
GROUP BY ref_student_id
";
                DataTable dt = qh.Select(qry);

                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["student_id"].ToString();
                        decimal hr;
                        if (decimal.TryParse(dr["hour_sum"].ToString(), out hr))
                        {
                            if (!srDict.ContainsKey(sid))
                            {
                                srDict.Add(sid, hr);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }

            foreach (StudentInfo si in StudentInfoList)
            {
                if (srDict.ContainsKey(si.StudentID))
                {
                    si.ServiceLearnScore = (int)(srDict[si.StudentID] / 2);
                    // 最高8分
                    if (si.ServiceLearnScore > 8)
                        si.ServiceLearnScore = 8;
                }

            }


            return StudentInfoList;
        }


        /// <summary>
        /// 取得成績冊用學生基本資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static List<rptStudentInfo> GetRptStudentInfoListByIDs(List<string> StudentIDList)
        {
            List<rptStudentInfo> StudentInfoList = new List<rptStudentInfo>();

            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string qry = @"
SELECT 
        student.id AS student_id
        ,class.class_name
        ,student.seat_no
        ,CASE student.gender WHEN '1' THEN '男' WHEN '0' THEN '女' ELSE '' END AS gender
        ,student.name AS student_name
        ,sems_history
        ,birthdate 
FROM student LEFT JOIN class 
ON student.ref_class_id = class.id 
WHERE student.status = 1 AND student.id IN(" + string.Join(",", StudentIDList.ToArray()) + @")
ORDER BY class.grade_year DESC,class.display_order,class.class_name,seat_no
";
                DataTable dt = qh.Select(qry);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        rptStudentInfo si = new rptStudentInfo();
                        si.StudentID = dr["student_id"].ToString();
                        si.ClassName = dr["class_name"].ToString();
                        si.SeatNo = dr["seat_no"].ToString();
                        si.Gender = dr["gender"].ToString();
                        si.SchoolYear = K12.Data.School.DefaultSchoolYear;
                        si.SchoolName = K12.Data.School.ChineseName;
                        si.Name = dr["student_name"].ToString();

                        DateTime dt1;

                        if (DateTime.TryParse(dr["birthdate"].ToString(), out dt1))
                        {
                            si.Birthday = dt1;
                        }

                        // 處理學期歷程
                        if (dr["sems_history"] != null)
                        {
                            try
                            {
                                XElement elms = XElement.Parse("<root>" + dr["sems_history"].ToString() + "</root>");
                                foreach (XElement elm in elms.Elements("History"))
                                {
                                    SemsHistoryInfo shi = new SemsHistoryInfo();
                                    if (elm.Attribute("SchoolYear") != null)
                                        shi.SchoolYear = elm.Attribute("SchoolYear").Value;

                                    if (elm.Attribute("Semester") != null)
                                        shi.Semester = elm.Attribute("Semester").Value;

                                    if (elm.Attribute("GradeYear") != null)
                                        shi.GradeYear = elm.Attribute("GradeYear").Value;

                                    string key = shi.SchoolYear + "_" + shi.Semester;

                                    string value = "";
                                    if (shi.GradeYear == "1" || shi.GradeYear == "7")
                                    {
                                        if (shi.Semester == "1")
                                            value = "七上";

                                        if (shi.Semester == "2")
                                            value = "七下";
                                    }

                                    if (shi.GradeYear == "2" || shi.GradeYear == "8")
                                    {
                                        if (shi.Semester == "1")
                                            value = "八上";

                                        if (shi.Semester == "2")
                                            value = "八下";
                                    }

                                    if (shi.GradeYear == "3" || shi.GradeYear == "9")
                                    {
                                        if (shi.Semester == "1")
                                            value = "九上";
                                    }

                                    if (value != "")
                                    {
                                        if (!si.SemsHistoryDict.ContainsKey(key))
                                            si.SemsHistoryDict.Add(key, value);
                                    }


                                    si.SemsHistoryInfoList.Add(shi);
                                }

                            }
                            catch (Exception ex) { }
                        }

                        StudentInfoList.Add(si);
                    }

                }

            }

            return StudentInfoList;
        }

        /// <summary>
        /// 取得領域資料並填入成績冊用
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="StudentInfoList"></param>
        /// <returns></returns>
        public static List<rptStudentInfo> FillRptDomainScoreInfo(List<string> StudentIDList, List<rptStudentInfo> StudentInfoList)
        {
            // 取得學生學期成績
            Dictionary<string, List<JHSemesterScoreRecord>> SemesterScoreRecordDict = new Dictionary<string, List<JHSemesterScoreRecord>>();
            List<JHSemesterScoreRecord> tmpSemsScore = JHSemesterScore.SelectByStudentIDs(StudentIDList);
            foreach (JHSemesterScoreRecord rec in tmpSemsScore)
            {
                if (!SemesterScoreRecordDict.ContainsKey(rec.RefStudentID))
                    SemesterScoreRecordDict.Add(rec.RefStudentID, new List<JHSemesterScoreRecord>());

                SemesterScoreRecordDict[rec.RefStudentID].Add(rec);
            }

            // 填入學期成績

            List<string> tmpDomainNameList = new List<string>();
            tmpDomainNameList.Add("健康體育");
            tmpDomainNameList.Add("健康與體育");
            tmpDomainNameList.Add("藝術人文");
            tmpDomainNameList.Add("藝術與人文");
            tmpDomainNameList.Add("綜合活動");

            foreach (rptStudentInfo si in StudentInfoList)
            {
                if (SemesterScoreRecordDict.ContainsKey(si.StudentID))
                {
                    foreach (JHSemesterScoreRecord semsRec in SemesterScoreRecordDict[si.StudentID])
                    {
                        string key = semsRec.SchoolYear + "_" + semsRec.Semester;

                        // 學期歷程有
                        if (si.SemsHistoryDict.ContainsKey(key))
                        {
                            foreach (string dname in semsRec.Domains.Keys)
                            {
                                if (tmpDomainNameList.Contains(dname))
                                {
                                    if (semsRec.Domains[dname].Score.HasValue)
                                    {
                                        if (!si.DomainScoreInfoDict.ContainsKey(dname))
                                        {
                                            rptDomainScoreInfo ds = new rptDomainScoreInfo();
                                            ds.StudentID = si.StudentID;
                                            ds.Name = dname;
                                            si.DomainScoreInfoDict.Add(dname, ds);
                                        }

                                        if (!si.DomainScoreInfoDict[dname].ScoreDict.ContainsKey(si.SemsHistoryDict[key]))
                                            si.DomainScoreInfoDict[dname].ScoreDict.Add(si.SemsHistoryDict[key], semsRec.Domains[dname].Score.Value);
                                    }

                                }
                            }
                        }
                    }
                }

                // 計算成績
                si.CalcDomainScoreInfoList();
            }

            return StudentInfoList;
        }

        /// <summary>
        /// 成績冊用體適能
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="StudentInfoList"></param>
        /// <returns></returns>
        public static List<rptStudentInfo> FillRptFitnessInfo(List<string> StudentIDList, List<rptStudentInfo> StudentInfoList)
        {
            if (StudentIDList.Count > 0)
            {
                Dictionary<string, List<rptFitnessInfo>> tmpFitDict = new Dictionary<string, List<rptFitnessInfo>>();

                try
                {
                    QueryHelper qh = new QueryHelper();
                    string qry = @"
SELECT 
    ref_student_id AS student_id
    ,test_date
    ,sit_and_reach
    ,sit_and_reach_degree
    ,standing_long_jump
    ,standing_long_jump_degree
    ,sit_up
    ,sit_up_degree
    ,cardiorespiratory
    ,cardiorespiratory_degree
FROM $ischool_student_fitness WHERE ref_student_id IN('" + string.Join("','", StudentIDList.ToArray()) + @"') 
ORDER BY test_date";
                    DataTable dt = qh.Select(qry);
                    if (dt != null)
                    {
                        foreach (DataRow dr in dt.Rows)
                        {
                            string sid = dr["student_id"].ToString();

                            if (!tmpFitDict.ContainsKey(sid))
                                tmpFitDict.Add(sid, new List<rptFitnessInfo>());

                            rptFitnessInfo fi = new rptFitnessInfo();
                            fi.StudentID = sid;

                            DateTime dt1;

                            if (DateTime.TryParse(dr["test_date"].ToString(), out dt1))
                            {
                                fi.TestDate = dt1;
                                fi.TestDateStr = (fi.TestDate.Year - 1911) + "/" + fi.TestDate.Month + "/" + fi.TestDate.Day;
                            }

                            fi.Sit_and_reach = dr["sit_and_reach"].ToString();
                            fi.Sit_and_reach_degree = dr["sit_and_reach_degree"].ToString();
                            fi.Standing_long_jump = dr["standing_long_jump"].ToString();
                            fi.Standing_long_jump_degree = dr["standing_long_jump_degree"].ToString();
                            fi.Sit_up = dr["sit_up"].ToString();
                            fi.Sit_up_degree = dr["sit_up_degree"].ToString();
                            fi.Cardiorespiratory = dr["cardiorespiratory"].ToString();
                            fi.Cardiorespiratory_degree = dr["cardiorespiratory_degree"].ToString();

                            tmpFitDict[sid].Add(fi);
                        }
                    }
                }
                catch (Exception ex)
                {

                }

                List<string> addStringList = new List<string>();
                addStringList.Add("金牌");
                addStringList.Add("銀牌");
                addStringList.Add("銅牌");

                // 填入學生體適能並年齡轉換
                foreach (rptStudentInfo si in StudentInfoList)
                {
                    if (tmpFitDict.ContainsKey(si.StudentID))
                    {
                        si.FitnessInfoList = tmpFitDict[si.StudentID];

                        // 轉換年齡,整理積分需要
                        foreach (rptFitnessInfo fi in si.FitnessInfoList)
                        {
                            fi.Age = si.GetAge(fi.TestDate);

                            int addCount = 0;

                            // sit_and_reach_degree
                            if (fi.Sit_and_reach_degree == "" || fi.Sit_and_reach_degree == "未檢測")
                            {

                            }
                            else
                            {
                                si.sit_and_reach_degreeList.Add(fi.Sit_and_reach_degree);
                                if (addStringList.Contains(fi.Sit_and_reach_degree))
                                    addCount++;
                            }

                            // standing_long_jump_degree
                            if (fi.Standing_long_jump_degree == "" || fi.Standing_long_jump_degree == "未檢測")
                            {

                            }
                            else
                            {
                                si.standing_long_jump_degreeList.Add(fi.Standing_long_jump_degree);
                                if (addStringList.Contains(fi.Standing_long_jump_degree))
                                    addCount++;
                            }


                            // sit_up_degree
                            if (fi.Sit_up_degree == "" || fi.Sit_up_degree == "未檢測")
                            {

                            }
                            else
                            {
                                si.sit_up_degreeList.Add(fi.Sit_up_degree);
                                if (addStringList.Contains(fi.Sit_up_degree))
                                    addCount++;
                            }

                            // cardiorespiratory_degree
                            if (fi.Cardiorespiratory_degree == "" || fi.Cardiorespiratory_degree == "未檢測")
                            {

                            }
                            else
                            {
                                si.cardiorespiratory_degreeList.Add(fi.Cardiorespiratory_degree);
                                if (addStringList.Contains(fi.Cardiorespiratory_degree))
                                    addCount++;
                            }

                            // 4 次都符合銅牌以上，符合加分
                            if (addCount == 4)
                                si.isAddFitnessScore = true;
                        }
                    }

                    si.CalcFitnessInfoList();
                }

            }
            return StudentInfoList;
        }

        /// <summary>
        /// 成績冊用獎懲
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="StudentInfoList"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static List<rptStudentInfo> FillRptMeritDemeritInfo(List<string> StudentIDList, List<rptStudentInfo> StudentInfoList, DateTime endDate)
        {
            List<JHDisciplineRecord> tmpRecordList = JHDiscipline.SelectByStudentIDs(StudentIDList);
            List<JHDisciplineRecord> DisciplineRecordList = new List<JHDisciplineRecord>();
            foreach (JHDisciplineRecord rec in tmpRecordList)
            {
                if (rec.OccurDate > endDate)
                    continue;
                DisciplineRecordList.Add(rec);
            }

            // 依日期排序
            DisciplineRecordList = DisciplineRecordList.OrderBy(x => x.OccurDate).ToList();

            Dictionary<string, List<JHDisciplineRecord>> DisciplineRecordDict = new Dictionary<string, List<JHDisciplineRecord>>();
            foreach (JHDisciplineRecord rec in DisciplineRecordList)
            {
                if (!DisciplineRecordDict.ContainsKey(rec.RefStudentID))
                    DisciplineRecordDict.Add(rec.RefStudentID, new List<JHDisciplineRecord>());

                DisciplineRecordDict[rec.RefStudentID].Add(rec);
            }

            foreach (rptStudentInfo si in StudentInfoList)
            {
                if (DisciplineRecordDict.ContainsKey(si.StudentID))
                {
                    foreach (JHDisciplineRecord rec in DisciplineRecordDict[si.StudentID])
                    {
                        string shKey = rec.SchoolYear + "_" + rec.Semester;

                        rptMeritDemeritInfo mdi = new rptMeritDemeritInfo();

                        if (si.SemsHistoryDict.ContainsKey(shKey))
                        {
                            mdi.Semester = si.SemsHistoryDict[shKey];
                        }
                        else
                            mdi.Semester = shKey;

                        mdi.StudentID = rec.RefStudentID;
                        mdi.OccurDate = (rec.OccurDate.Year - 1911) + "/" + rec.OccurDate.Month + "/" + rec.OccurDate.Day;
                        mdi.Reason = rec.Reason;
                        if (rec.MeritA.HasValue)
                            mdi.MA = rec.MeritA.Value;

                        if (rec.MeritB.HasValue)
                            mdi.MB = rec.MeritB.Value;

                        if (rec.MeritC.HasValue)
                            mdi.MC = rec.MeritC.Value;

                        mdi.Cleand = rec.Cleared;
                        
                        if (rec.DemeritA.HasValue)
                            mdi.DA = rec.DemeritA.Value;

                        if (rec.DemeritB.HasValue)
                            mdi.DB = rec.DemeritB.Value;

                        if (rec.DemeritC.HasValue)
                            mdi.DC = rec.DemeritC.Value;

                        si.MeritDemeritInfoList.Add(mdi);
                    }

                }

                // 計算獎懲統計
                si.CalcMeritDemeritInfoList();
            }

            return StudentInfoList;
        }


        /// <summary>
        /// 成績冊用服務學習
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="StudentInfoList"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static List<rptStudentInfo> FillRptServiceInfo(List<string> StudentIDList, List<rptStudentInfo> StudentInfoList, DateTime endDate)
        {

            if (StudentIDList.Count > 0)
            {
                endDate = endDate.AddDays(1);
                string strEndDate = endDate.Year + "-" + endDate.Month + "-" + endDate.Day;

                Dictionary<string, List<rptServiceInfo>> tmpSrvDict = new Dictionary<string, List<rptServiceInfo>>();

                QueryHelper qh = new QueryHelper();
                string qry = @"
SELECT 
    ref_student_id
    ,occur_date
    ,internal_or_external
    ,hours
    ,reason
    ,organizers 
FROM
 $k12.service.learning.record
WHERE occur_date < '" + strEndDate + "' AND ref_student_id IN('" + string.Join("','", StudentIDList.ToArray()) + @"')
ORDER BY ref_student_id,occur_date";

                DataTable dt = qh.Select(qry);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        rptServiceInfo sif = new rptServiceInfo();
                        sif.StudentID = dr["ref_student_id"].ToString();
                        sif.InternalOrExternal = dr["internal_or_external"].ToString();

                        DateTime ddt;
                        if (DateTime.TryParse(dr["occur_date"].ToString(), out ddt))
                        {
                            sif.OccurDate = (ddt.Year - 1911) + "/" + ddt.Month + "/" + ddt.Day;
                        }

                        decimal hr;
                        sif.Hours = 0;
                        if (decimal.TryParse(dr["hours"].ToString(), out hr))
                            sif.Hours = hr;

                        sif.Reason = dr["reason"].ToString();
                        sif.Organizers = dr["organizers"].ToString();

                        if (!tmpSrvDict.ContainsKey(sif.StudentID))
                            tmpSrvDict.Add(sif.StudentID, new List<rptServiceInfo>());

                        tmpSrvDict[sif.StudentID].Add(sif);
                    }
                }

                foreach (rptStudentInfo si in StudentInfoList)
                {
                    if (tmpSrvDict.ContainsKey(si.StudentID))
                    {
                        si.ServiceInfoList = tmpSrvDict[si.StudentID];
                        si.CalcServiceInfoList();
                    }
                }

            }

            return StudentInfoList;
        }

        /// <summary>
        /// 成績冊使用 低收入戶
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="StudentInfoList"></param>
        /// <returns></returns>
        public static List<rptStudentInfo> FillrptIncomeType(List<string> StudentIDList, List<rptStudentInfo> StudentInfoList)
        {
            try
            {
                // 取得低收入資料
                QueryHelper qh = new QueryHelper();

                Dictionary<string, DataRow> compDict = new Dictionary<string, DataRow>();
                string qry = @"
SELECT ref_student_id AS student_id
,income_type 
FROM student_info_ext WHERE ref_student_id IN(" + string.Join(",", StudentIDList.ToArray()) + @")

";
                DataTable dt = qh.Select(qry);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["student_id"].ToString();
                        if (!compDict.ContainsKey(sid))
                            compDict.Add(sid, dr);
                    }
                }

                // 填入資料
                foreach (rptStudentInfo si in StudentInfoList)
                {
                    si.IncomeType1 = false;
                    if (compDict.ContainsKey(si.StudentID))
                    {
                        if (compDict[si.StudentID]["income_type"] != null)
                        {
                            if (compDict[si.StudentID]["income_type"].ToString() == "低")
                                si.IncomeType1 = true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // throw ex;
            }

            return StudentInfoList;
        }




    }
}
