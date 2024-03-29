﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using FISCA.Permission;
using FISCA.Presentation;

namespace ChiaYiExcessCompetition
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {

            #region 嘉義免試入學-免試入學成績冊(超額比序均衡學習成績證明單)
            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["學生", "資料統計"];
            rbItem1["報表"]["嘉義免試入學"]["超額比序均衡學習成績證明單"].Enable = UserAcl.Current["7EE379CC-41BC-4A31-A403-232F0D0F6EB0"].Executable;
            rbItem1["報表"]["嘉義免試入學"]["超額比序均衡學習成績證明單"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    ScoreReportForm srf = new ScoreReportForm();
                    srf.SetStudentIDs(K12.Presentation.NLDPanels.Student.SelectedSource);
                    srf.ShowDialog();

                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選學生");
                    return;
                }
            };

            // 嘉義免試入學-免試入學成績冊//超額比序均衡學習成績證明單
            Catalog catalog1 = RoleAclSource.Instance["學生"]["功能按鈕"];
            catalog1.Add(new RibbonFeature("7EE379CC-41BC-4A31-A403-232F0D0F6EB0", "嘉義免試入學-超額比序均衡學習成績證明單"));

            #endregion


            #region 送審用匯入檔
            RibbonBarItem rbItem2 = MotherForm.RibbonBarItems["教務作業", "資料統計"];
            rbItem2["報表"]["嘉義免試入學"]["送審用匯入檔"].Enable = UserAcl.Current["15549571-41EC-412A-9529-66D63A4DC8FE"].Executable;
            rbItem2["報表"]["嘉義免試入學"]["送審用匯入檔"].Click += delegate
            {
                SubmitForReviewForm sfr = new SubmitForReviewForm();
                sfr.ShowDialog();

            };

            // 嘉義免試入學-送審用匯入檔
            Catalog catalog2 = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            catalog2.Add(new RibbonFeature("15549571-41EC-412A-9529-66D63A4DC8FE", "嘉義免試入學-送審用匯入檔"));

            #endregion


        }
    }
}
