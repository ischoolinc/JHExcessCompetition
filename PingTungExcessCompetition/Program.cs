using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using FISCA.Permission;
using FISCA.Presentation;

namespace PingTungExcessCompetition
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {

            #region 屏東免試入學-班級服務表現
            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["班級", "資料統計"];
            rbItem1["報表"]["屏東免試入學"]["班級服務表現"].Enable = UserAcl.Current["CDBF4D69-AD83-46AB-8F2B-7DEE4ADD03AE"].Executable;
            rbItem1["報表"]["屏東免試入學"]["班級服務表現"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {

                    ServiceReportForm srf = new ServiceReportForm();
                    srf.SetClassIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
                    srf.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選班級");
                    return;
                }

            };

            // 屏東免試入學-班級服務表現
            Catalog catalog1 = RoleAclSource.Instance["班級"]["功能按鈕"];
            catalog1.Add(new RibbonFeature("CDBF4D69-AD83-46AB-8F2B-7DEE4ADD03AE", "屏東免試入學-班級服務表現"));

            #endregion


            #region 送審用匯入檔
            RibbonBarItem rbItem2 = MotherForm.RibbonBarItems["教務作業", "資料統計"];
            rbItem2["報表"]["屏東免試入學"]["送審用匯入檔"].Enable = UserAcl.Current["D67C5AB8-4DBB-413A-9688-DD9F0075B479"].Executable;
            rbItem2["報表"]["屏東免試入學"]["送審用匯入檔"].Click += delegate
            {
                SubmitForReviewForm sfr = new SubmitForReviewForm();
                sfr.ShowDialog();

            };

            // 屏東免試入學-送審用匯入檔
            Catalog catalog2 = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            catalog2.Add(new RibbonFeature("D67C5AB8-4DBB-413A-9688-DD9F0075B479", "屏東免試入學-送審用匯入檔"));

            #endregion

            #region 設定幹部限制
            RibbonBarItem rbItem3 = MotherForm.RibbonBarItems["教務作業", "資料統計"];
            rbItem2["報表"]["屏東免試入學"]["設定幹部限制"].Enable = UserAcl.Current["8443B0A6-A7B6-478B-9120-63C4F7B6AE14"].Executable;
            rbItem2["報表"]["屏東免試入學"]["設定幹部限制"].Click += delegate
            {
                setCadreNameForm scnf = new setCadreNameForm();
                scnf.ShowDialog();

            };

            // 屏東免試入學-設定幹部限制
            Catalog catalog3 = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            catalog2.Add(new RibbonFeature("8443B0A6-A7B6-478B-9120-63C4F7B6AE14", "屏東免試入學-設定幹部限制"));

            #endregion

        }
    }
}
