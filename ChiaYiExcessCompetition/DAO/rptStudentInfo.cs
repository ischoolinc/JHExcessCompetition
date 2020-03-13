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
        /// 領域成績
        /// </summary>
        public List<rptDomainScoreInfo> DomainScoreInfoList = new List<rptDomainScoreInfo>();

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
        /// 均衡學習_健康與體育_平均
        /// </summary>
        public decimal DomainHelAvgScore = 0;

        /// <summary>
        /// 均衡學習_藝術與人文_平均
        /// </summary>
        public decimal DomainArtAvgScore = 0;

        /// <summary>
        /// 均衡學習_綜合活動_平均
        /// </summary>
        public decimal DomainActAvgScore = 0;

        /// <summary>
        /// 均衡學習_積分
        /// </summary>
        public int DomainIScore = 0;

        /// <summary>
        /// 品德表現_獎懲_大功統計
        /// </summary>
        public int MACount = 0;

        /// <summary>
        /// 品德表現_獎懲_小功統計
        /// </summary>
        public int MBCount = 0;

        /// <summary>
        /// 品德表現_獎懲_嘉獎統計
        /// </summary>
        public int MCCount = 0;

        /// <summary>
        /// 品德表現_獎懲_大過統計
        /// </summary>
        public int DACount = 0;

        /// <summary>
        /// 品德表現_獎懲_小過統計
        /// </summary>
        public int DBCount = 0;

        /// <summary>
        /// 品德表現_獎懲_警告統計
        /// </summary>
        public int DCCount = 0;

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
        public int ServiceOutHourCount = 0;

        /// <summary>
        /// 品德表現_服務學習_校內時數統計
        /// </summary>
        public int ServiceInHourCount = 0;

        /// <summary>
        /// 品德表現_體適能_積分
        /// </summary>
        public int FitnessIScore = 0;

        public void CalcDomainScoreInfoList() { }

        public void CalcMeritDemeritInfoList() { }

        public void CalcFitnessInfoList() { }

        public void CalcServiceInfoList() { }

    }
}
