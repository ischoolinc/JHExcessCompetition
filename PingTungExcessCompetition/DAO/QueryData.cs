﻿using System;
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
                        si.CalcScore();


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

    }
}