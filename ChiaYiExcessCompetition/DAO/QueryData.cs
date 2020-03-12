using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Data;
using System.Data;
using System.Xml.Linq;

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

            return StudentInfoList;
        }


    }
}
