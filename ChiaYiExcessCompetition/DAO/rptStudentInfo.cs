using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChiaYiExcessCompetition.DAO
{
    /// <summary>
    /// 成績冊用學生資訊
    /// </summary>
    public class rptStudentInfo
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 學年度
        /// </summary>
        public string SchoolYear { get; set; }

        /// <summary>
        /// 學校名稱
        /// </summary>
        public string SchoolName { get; set; }

        /// <summary>
        /// 班級
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        public string SeatNo { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 性別
        /// </summary>
        public string Gender { get; set; }

        /// <summary>
        /// 學習歷程
        /// </summary>
        public List<SemsHistoryInfo> SemsHistoryInfoList = new List<SemsHistoryInfo>();

        public Dictionary<string, string> SemsHistoryDict = new Dictionary<string, string>();

        /// <summary>
        /// 領域成績
        /// </summary>
        public Dictionary<string, rptDomainScoreInfo> DomainScoreInfoDict = new Dictionary<string, rptDomainScoreInfo>();

        /// <summary>
        /// 獎懲資料
        /// </summary>
        public List<rptMeritDemeritInfo> MeritDemeritInfoList = new List<rptMeritDemeritInfo>();

        /// <summary>
        /// 體適能資料
        /// </summary>
        public List<rptFitnessInfo> FitnessInfoList = new List<rptFitnessInfo>();

        /// <summary>
        /// 服務學習資料
        /// </summary>
        public List<rptServiceInfo> ServiceInfoList = new List<rptServiceInfo>();


        /// <summary>
        /// 均衡學習_積分
        /// </summary>
        public int DomainIScore = 0;

        /// <summary>
        /// 品德表現_獎懲_大功統計
        /// </summary>
        public int MASum = 0;

        /// <summary>
        /// 品德表現_獎懲_小功統計
        /// </summary>
        public int MBSum = 0;

        /// <summary>
        /// 品德表現_獎懲_嘉獎統計
        /// </summary>
        public int MCSum = 0;

        /// <summary>
        /// 品德表現_獎懲_大過統計
        /// </summary>
        public int DASum = 0;

        /// <summary>
        /// 品德表現_獎懲_小過統計
        /// </summary>
        public int DBSum = 0;

        /// <summary>
        /// 品德表現_獎懲_警告統計
        /// </summary>
        public int DCSum = 0;

        /// <summary>
        /// 品德表現_獎懲_銷過統計
        /// </summary>
        public int MDCleanCount = 0;

        /// <summary>
        /// 品德表現_獎懲_積分
        /// </summary>
        public int MDIScore = 0;

        /// <summary>
        /// 品德表現_服務學習_積分
        /// </summary>
        public int ServiceIScore = 0;

        /// <summary>
        /// 品德表現_服務學習_校外時數統計
        /// </summary>
        public decimal ServiceOutHourCount = 0;

        /// <summary>
        /// 品德表現_服務學習_校內時數統計
        /// </summary>
        public decimal ServiceInHourCount = 0;

        /// <summary>
        /// 品德表現_體適能_積分
        /// </summary>
        public int FitnessIScore = 0;

        public void CalcDomainScoreInfoList()
        {
            foreach (string dname in DomainScoreInfoDict.Keys)
            {
                decimal score = 0;
                foreach (string na in DomainScoreInfoDict[dname].ScoreDict.Keys)
                {
                    score += DomainScoreInfoDict[dname].ScoreDict[na];
                }

                // 計算平均
                DomainScoreInfoDict[dname].AvgScore = Math.Round(score / 5, 0, MidpointRounding.AwayFromZero);

                // 計算積分
                if (DomainScoreInfoDict[dname].AvgScore >= 60)
                    DomainIScore += 3;
            }
        }

        public void CalcMeritDemeritInfoList()
        {
            foreach (rptMeritDemeritInfo rec in MeritDemeritInfoList)
            {
                if (rec.Cleand != "是")
                {
                    MASum += rec.MA;
                    MBSum += rec.MB;
                    MCSum += rec.MC;
                    DASum += rec.DA;
                    DBSum += rec.DB;
                    DCSum += rec.DC;
                }

                if (rec.Cleand == "是")
                    MDCleanCount += 1;
            }

            int MMSum = 0, DDSum = 0;
            // 計算積分，功過相抵
            foreach (rptMeritDemeritInfo rec in MeritDemeritInfoList)
            {
                if (rec.Cleand == "是")
                    continue;

                // 功
                MMSum += (rec.MA * 9 + rec.MB * 3 + rec.MC);
                // 過
                DDSum += (rec.DA * 9 + rec.DB * 3 + rec.DC);
            }
            int MDSum = MMSum - DDSum;

            // 預設6分
            MDIScore = 6;

            if (MMSum >= 0)
                MDIScore += MMSum;
            else if (MMSum > -9)
            {
                MDIScore = 3;
            }
            else
            {
                MDIScore = 0;
            }

            // 最高 12 分
            if (MDIScore > 12)
                MDIScore = 12;


        }

        public void CalcFitnessInfoList() { }

        public void CalcServiceInfoList()
        {

            decimal sum = 0;
            foreach (rptServiceInfo sif in ServiceInfoList)
            {
                if (sif.InternalOrExternal == "校內")
                {
                    ServiceInHourCount += sif.Hours;
                }

                if (sif.InternalOrExternal == "校外")
                {
                    ServiceInHourCount += sif.Hours;
                }

                sum += sif.Hours;
            }

            ServiceIScore = (int)(sum / 2);

            // 最高8分
            if (ServiceIScore > 8)
                ServiceIScore = 8;
        }

    }
}
