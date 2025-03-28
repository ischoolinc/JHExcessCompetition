﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChiaYiExcessCompetition.DAO
{
    /// <summary>
    /// 成績冊用體適能
    /// </summary>
    public class rptFitnessInfo
    {
        public string StudentID { get; set; }

        /// <summary>
        /// 品德表現_體適能_檢測日期
        /// </summary>
        public string TestDateStr { get; set; }

        /// <summary>
        /// 測驗日期，計算年齡使用
        /// </summary>
        public DateTime TestDate { get; set; }

        /// <summary>
        /// 品德表現_體適能_年齡
        /// </summary>
        public int Age { get; set; }

        /// <summary>
        /// 品德表現_體適能_坐姿體前彎_成績
        /// </summary>
        public string Sit_and_reach { get; set; }

        /// <summary>
        /// 品德表現_體適能_坐姿體前彎_等級
        /// </summary>
        public string Sit_and_reach_degree { get; set; }

        /// <summary>
        /// 品德表現_體適能_立定跳遠_成績
        /// </summary>
        public string Standing_long_jump { get; set; }

        /// <summary>
        /// 品德表現_體適能_立定跳遠_等級
        /// </summary>
        public string Standing_long_jump_degree { get; set; }

        /// <summary>
        /// 品德表現_體適能_仰臥起坐_成績
        /// </summary>
        public string Sit_up { get; set; }

        /// <summary>
        /// 品德表現_體適能_仰臥起坐_等級
        /// </summary>
        public string Sit_up_degree { get; set; }

        /// <summary>
        /// 品德表現_體適能_公尺跑走_成績
        /// </summary>
        public string Cardiorespiratory { get; set; }

        /// <summary>
        /// 品德表現_體適能_公尺跑走_等級
        /// </summary>
        public string Cardiorespiratory_degree { get; set; }

        /// <summary>
        ///  仰臥捲腹
        /// </summary>        
        public string Curl { get; set; }

        /// <summary>
        ///  仰臥捲腹常模
        /// </summary>        
        public string CurlDegree { get; set; }

        /// <summary>
        ///  漸速耐力跑
        /// </summary>        
        public string Pacer { get; set; }

        /// <summary>
        ///  漸速耐力跑常模
        /// </summary>        
        public string PacerDegree { get; set; }
    }
}
