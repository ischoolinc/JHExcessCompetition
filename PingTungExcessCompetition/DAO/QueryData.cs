using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using System.Data;
using System.Xml.Linq;


namespace PingTungExcessCompetition.DAO
{
    public class QueryData
    {
        public static Dictionary<string, string> GetClassTeacherNameDictByClassID(List<string> ClassIDList)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            if (ClassIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "SELECT " +
                    "class.id AS class_id" +
                    ",teacher.teacher_name" +
                    ",teacher.nickname " +
                    "FROM " +
                    "class LEFT JOIN teacher " +
                    "ON class.ref_teacher_id = teacher.id " +
                    "WHERE teacher.status = 1 AND class.id IN(" + string.Join(",", ClassIDList.ToArray()) + ") " +
                    "ORDER BY class.grade_year,class.display_order,class_name";

                DataTable dt = qh.Select(query);

                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string class_id = dr["class_id"].ToString();
                        string teacher_name = "", nickname = "";
                        if (dr["teacher_name"] != null)
                            teacher_name = dr["teacher_name"].ToString();

                        if (dr["nickname"] != null)
                        {
                            nickname = dr["nickname"].ToString();
                        }

                        if (!value.ContainsKey(class_id))
                        {
                            if (!string.IsNullOrWhiteSpace(nickname))
                            {
                                teacher_name = teacher_name + "(" + nickname + ")";
                            }
                            value.Add(class_id, teacher_name);
                        }
                    }
                }

            }
            return value;
        }


        /// <summary>
        /// 透過班級編號 取得 班級學生一般狀態
        /// </summary>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, List<StudentInfo>> GetClassStudentDict(List<string> ClassIDList)
        {
            Dictionary<string, List<StudentInfo>> value = new Dictionary<string, List<StudentInfo>>();
            try
            {

                string classIDs = string.Join(",", ClassIDList.ToArray());

                string qryStud = @"
SELECT student.ref_class_id AS class_id
,student.id AS student_id
,seat_no,student.name AS student_name
,sems_history 
FROM student 
WHERE 
student.status=1 AND 
student.ref_class_id IN(" + classIDs + @") 
ORDER BY student.ref_class_id,seat_no";

                QueryHelper qh = new QueryHelper();
                DataTable dtStud = qh.Select(qryStud);

                // 取得班級異動資料
                string qryUpdateRec = @"
SELECT student.id AS student_id
,update_record.school_year
,update_record.semester
,CASE
        WHEN update_code = '2' THEN '畢業' 
        WHEN update_code = '3' THEN '轉入' 
        WHEN update_code = '4' THEN '轉出' 
        WHEN update_code = '5' THEN '休學' 
        WHEN update_code = '6' THEN '復學' 
        WHEN update_code = '7' THEN '中輟' 
        WHEN update_code = '8' THEN '續讀' 
        WHEN update_code = '9' THEN '更正學籍' 
        WHEN update_code = '10' THEN '延長修業年限' 
        WHEN update_code = '11' THEN '死亡' 
        ELSE ''
    END 
 AS update_desc 
FROM update_record 
INNER JOIN student
ON update_record.ref_student_id = student.id 
WHERE update_code<>'1' AND student.status=1 AND student.ref_class_id IN(" + classIDs + @") ORDER BY student.id,update_date DESC
";

                DataTable dtUpdRec = qh.Select(qryUpdateRec);
                Dictionary<string, List<string>> studUpdaeRecDict = new Dictionary<string, List<string>>();

                if (dtUpdRec != null)
                {
                    foreach (DataRow dr in dtUpdRec.Rows)
                    {
                        string sid = dr["student_id"].ToString();
                        if (!studUpdaeRecDict.ContainsKey(sid))
                            studUpdaeRecDict.Add(sid, new List<string>());

                        string msg = dr["school_year"].ToString() + " " + dr["semester"].ToString() + " " + dr["update_desc"].ToString();

                        studUpdaeRecDict[sid].Add(msg);
                    }
                }


                // 取得幹部資料
                Dictionary<string, List<CadreInfo>> studCadDict = new Dictionary<string, List<CadreInfo>>();
                string qryCadre = @"
SELECT 
        student.id AS student_id
        ,schoolyear AS school_year
        ,semester
        ,referencetype
        ,cadrename
 FROM 
    student INNER JOIN $behavior.thecadre
         ON student.id = CAST($behavior.thecadre.studentid AS INTEGER) 
  WHERE student.status=1 AND student.ref_class_id IN(" + classIDs + @") 
     ORDER BY student.id,$behavior.thecadre.schoolyear,semester
";

                DataTable dtCadre = qh.Select(qryCadre);
                if (dtCadre != null)
                {
                    foreach (DataRow dr in dtCadre.Rows)
                    {
                        string sid = dr["student_id"].ToString();

                        if (!studCadDict.ContainsKey(sid))
                            studCadDict.Add(sid, new List<CadreInfo>());

                        CadreInfo ci = new CadreInfo();
                        ci.StudentID = sid;
                        ci.SchoolYear = dr["school_year"].ToString();
                        ci.Semester = dr["semester"].ToString();
                        ci.ReferenceType = dr["referencetype"].ToString();
                        ci.CadreName = dr["cadrename"].ToString();
                        studCadDict[sid].Add(ci);
                    }
                }


                // 班級幹部限制
                List<string> CadreName1 = Global.GetCadreName1();

                // 學生資料取得與整理需要相關資料
                if (dtStud != null)
                {
                    foreach (DataRow dr in dtStud.Rows)
                    {
                        StudentInfo si = new StudentInfo();
                        if (dr["seat_no"] != null)
                            si.SeatNo = dr["seat_no"].ToString();
                        else
                            si.SeatNo = "";
                        si.StudentID = dr["student_id"].ToString();
                        si.ClassID = dr["class_id"].ToString();
                        si.StudentName = dr["student_name"].ToString();

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

                        // 放入異動資料
                        if (studUpdaeRecDict.ContainsKey(si.StudentID))
                        {
                            si.ServiceMemo = studUpdaeRecDict[si.StudentID];
                        }

                        // 放入幹部資料
                        if (studCadDict.ContainsKey(si.StudentID))
                        {
                            si.CadreInfoList = studCadDict[si.StudentID];
                        }

                        // 計算幹部積分
                        si.CalcCadreScore(CadreName1);


                        if (!value.ContainsKey(si.ClassID))
                            value.Add(si.ClassID, new List<StudentInfo>());

                        value[si.ClassID].Add(si);
                    }

                }


            }
            catch (Exception ex)
            {
                throw ex;
            }

            return value;
        }


        /// <summary>
        /// 取得學生所有可選類別 return idprefix:name
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
                        value.Add(id, tagName);
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
        /// 有取得語言認證學生ID
        /// </summary>
        /// <param name="studIDList"></param>
        /// <returns></returns>
        public static List<string> GetLanguageCertificate(List<string> studIDList)
        {
            List<string> value = new List<string>();
            if (studIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string qry = @"
SELECT refid AS student_id 
FROM 
$stud.userdefinedata 
WHERE fieldname = '語言認證' AND value = '是' AND refid IN('" + string.Join("','", studIDList.ToArray()) + "') ";

                DataTable dt = qh.Select(qry);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["student_id"].ToString();
                        if (!value.Contains(sid))
                            value.Add(sid);
                    }
                }
            }

            return value;
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
                // 取得特定日期前體適能資料
                QueryHelper qh = new QueryHelper();
                endDate = endDate.AddDays(1);
                string strEndDate = endDate.Year + "-" + endDate.Month + "-" + endDate.Day;
                Dictionary<string, List<DataRow>> finDict = new Dictionary<string, List<DataRow>>();
                string qry = @"
SELECT 
ref_student_id
,test_date
,sit_and_reach_degree
,standing_long_jump_degree
,sit_up_degree
,cardiorespiratory_degree
FROM $ischool_student_fitness WHERE test_date <'" + strEndDate + @"' AND ref_student_id IN('" + string.Join("','", StudentIDList.ToArray()) + @"')
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
                        foreach (DataRow dr in finDict[si.StudentID])
                        {
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
                                }
                            }
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
        /// 取得競賽資料並填入
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="StudentInfoList"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static List<StudentInfo> FillStudentCompetitionScore(List<string> StudentIDList, List<StudentInfo> StudentInfoList, DateTime endDate)
        {
            try
            {
                // 取得特定日期前競賽資料
                QueryHelper qh = new QueryHelper();
                endDate = endDate.AddDays(1);
                string strEndDate = endDate.Year + "-" + endDate.Month + "-" + endDate.Day;
                Dictionary<string, List<DataRow>> compDict = new Dictionary<string, List<DataRow>>();
                string qry = @"
SELECT 
    ref_student_id
    ,school_year
    ,habitude
    ,organizer
    ,setting_name
    ,max(bt_integral)  bt_integral
FROM $competition.performance.student INNER JOIN $competition.performance.rank
ON $competition.performance.student.rank_name = $competition.performance.rank.bt_rank
WHERE ref_student_id IN ('" + string.Join("','", StudentIDList.ToArray()) + "') AND certificate_date < '" + strEndDate + @"'
GROUP BY ref_student_id
,school_year
,habitude
,organizer
,setting_name

";
                DataTable dt = qh.Select(qry);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["ref_student_id"].ToString();
                        if (!compDict.ContainsKey(sid))
                            compDict.Add(sid, new List<DataRow>());

                        compDict[sid].Add(dr);
                    }
                }

                // 填入資料
                foreach (StudentInfo si in StudentInfoList)
                {
                    if (compDict.ContainsKey(si.StudentID))
                    {
                        foreach (DataRow dr in compDict[si.StudentID])
                        {
                            int sc = 0;
                            if (int.TryParse(dr["bt_integral"].ToString(), out sc))
                            {
                                si.CompetitionScoreD.Add(sc);
                            }
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
        /// 取得幹部資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="StudentInfoList"></param>
        /// <returns></returns>
        public static List<StudentInfo> FillCad(List<string> StudentIDList, List<StudentInfo> StudentInfoList)
        {
            try
            {
                // 取得中低收入資料
                QueryHelper qh = new QueryHelper();

                Dictionary<string, List<CadreInfo>> studCadDict = new Dictionary<string, List<CadreInfo>>();
                string qry = @"
SELECT 
        student.id AS student_id
        ,schoolyear AS school_year
        ,semester
        ,referencetype
        ,cadrename
 FROM 
    student INNER JOIN $behavior.thecadre
         ON student.id = CAST($behavior.thecadre.studentid AS INTEGER) 
  WHERE student.status=1 AND student.id IN(" + string.Join(",", StudentIDList.ToArray()) + @") 
     ORDER BY student.id,$behavior.thecadre.schoolyear,semester
";
                DataTable dt = qh.Select(qry);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["student_id"].ToString();
                        if (!studCadDict.ContainsKey(sid))
                        {
                            studCadDict.Add(sid, new List<CadreInfo>());
                        }

                        CadreInfo ci = new CadreInfo();
                        ci.SchoolYear = dr["school_year"].ToString();
                        ci.Semester = dr["semester"].ToString();
                        ci.StudentID = dr["student_id"].ToString();
                        ci.CadreName = dr["cadrename"].ToString();
                        ci.ReferenceType = dr["referencetype"].ToString();
                        studCadDict[sid].Add(ci);
                    }
                }

                // 填入資料
                foreach (StudentInfo si in StudentInfoList)
                {
                    if (studCadDict.ContainsKey(si.StudentID))
                    {
                        si.CadreInfoList = studCadDict[si.StudentID];
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
