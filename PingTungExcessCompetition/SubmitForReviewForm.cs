using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using Aspose.Cells;
using System.IO;
using FISCA.UDT;
using System.Xml.Linq;

namespace PingTungExcessCompetition
{
    public partial class SubmitForReviewForm : BaseForm
    {
        BackgroundWorker bgWorkerExport = new BackgroundWorker();
        public Configure _Configure { get; private set; }
        AccessHelper _accessHelper = new AccessHelper();
    

        public SubmitForReviewForm()
        {
            InitializeComponent();
            bgWorkerExport.DoWork += BgWorkerExport_DoWork;
            bgWorkerExport.RunWorkerCompleted += BgWorkerExport_RunWorkerCompleted;
            bgWorkerExport.ProgressChanged += BgWorkerExport_ProgressChanged;
            bgWorkerExport.WorkerReportsProgress = true;

        }

        private void BgWorkerExport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生中 ... ", e.ProgressPercentage);
        }

        private void BgWorkerExport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                try
                {
                    Workbook wb = (Workbook)e.Result;
                    if (wb != null)
                    {
                        string reportName = "屏東免試入學-送審用匯入檔";

                        string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                        if (!Directory.Exists(path))
                            Directory.CreateDirectory(path);
                        path = Path.Combine(path, reportName + ".xlsx");

                        if (File.Exists(path))
                        {
                            int i = 1;
                            while (true)
                            {
                                string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                                if (!File.Exists(newPath))
                                {
                                    path = newPath;
                                    break;
                                }
                            }
                        }

                        Workbook wbNew = new Workbook();
                        try
                        {

                            wbNew = wb;
                            wbNew.Save(path, SaveFormat.Xlsx);
                            System.Diagnostics.Process.Start(path);
                        }
                        catch
                        {
                            System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                            sd.Title = "另存新檔";
                            sd.FileName = reportName + ".xlsx";
                            sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
                            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                try
                                {
                                    wbNew.Save(sd.FileName, SaveFormat.Xlsx);

                                }
                                catch
                                {
                                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                    return;
                                }
                            }
                        }
                    }

                    dtDate.Enabled = btnExport.Enabled = true;
                }
                catch (Exception ex)
                {
                    MsgBox.Show("發生錯誤：" + ex.Message);
                }
            }
            else
            {
                MsgBox.Show("發生錯誤：" + e.Error.Message);
            }
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        private void BgWorkerExport_DoWork(object sender, DoWorkEventArgs e)
        {
            bgWorkerExport.ReportProgress(1);
            try
            {
                // 取得預設樣板
                Workbook wb = new Workbook(new MemoryStream(Properties.Resources.Template));


                bgWorkerExport.ReportProgress(100);

                e.Result = wb;


            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            // 取得畫面上日期
            _Configure.EndDate = dtDate.Value;
            dtDate.Enabled = false;
            btnExport.Enabled = false;
            bgWorkerExport.RunWorkerAsync();

        }

        private void SubmitForReviewForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            LoadTemplate();
            dtDate.Value = _Configure.EndDate;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveConfig();
        }

        /// <summary>
        /// 儲存設定檔
        /// </summary>
        private void SaveConfig()
        {

        }

        private void LoadTemplate()
        {
            try
            {
                UserControlEnable(false);
                List<Configure> confList = _accessHelper.Select<Configure>();
                if (confList != null && confList.Count > 0)
                {
                    _Configure = confList[0];
                    if (_Configure.EndDate == null)
                        _Configure.EndDate = new DateTime(DateTime.Now.Year, 4, 30);
                    else
                    {
                        if (_Configure.EndDate.Date.Year < DateTime.Now.Year)
                            _Configure.EndDate = new DateTime(DateTime.Now.Year, 4, 30);
                    }

                    if (string.IsNullOrEmpty(_Configure.MappingContent))
                    {
                        _Configure.MappingContent = DefaultConfigXML().ToString();
                    }
                }
                else
                {
                    _Configure = new Configure();
                    _Configure.Name = "屏東免試入學-班級服務表現";
                    if (_Configure.EndDate == null)
                        _Configure.EndDate = new DateTime(DateTime.Now.Year, 4, 30);

                    if (string.IsNullOrEmpty(_Configure.MappingContent))
                    {
                        _Configure.MappingContent = DefaultConfigXML().ToString();
                    }
                }

                _Configure.Save();

                UserControlEnable(true);
            }
            catch (Exception ex)
            {
                MsgBox.Show("讀取樣板發生錯誤，" + ex.Message);
            }

        }


        /// <summary>
        /// 預設對照XML
        /// </summary>
        /// <returns></returns>
        private XElement DefaultConfigXML()
        {
            XElement elmRoot = new XElement("MappingConfig");
            List<string> gList = new List<string>();
            gList.Add("學生身分");
            gList.Add("身心障礙");
            gList.Add("學生報名身分設定");
            gList.Add("失業勞工子女");

            // 0 一般生
            Dictionary<string, string> g1ItemList = new Dictionary<string, string>();
            g1ItemList.Add("原住民", "1");
            g1ItemList.Add("派外人員子女", "2");
            g1ItemList.Add("蒙藏生", "3");
            g1ItemList.Add("回國僑生", "4");
            g1ItemList.Add("港澳生", "5");
            g1ItemList.Add("退伍軍人", "6");
            g1ItemList.Add("境外優秀科學技術人才子女", "7");

            // 0	無
            Dictionary<string, string> g2ItemList = new Dictionary<string, string>();
            g2ItemList.Add("智能障礙", "1");
            g2ItemList.Add("視覺障礙", "2");
            g2ItemList.Add("聽覺障礙", "3");
            g2ItemList.Add("語言障礙", "4");
            g2ItemList.Add("肢體障礙", "5");
            g2ItemList.Add("腦性麻痺", "6");
            g2ItemList.Add("身體病弱", "7");
            g2ItemList.Add("情緒行為障礙", "8");
            g2ItemList.Add("學習障礙", "9");
            g2ItemList.Add("多重障礙", "A");
            g2ItemList.Add("自閉症", "B");
            g2ItemList.Add("發展遲緩", "C");
            g2ItemList.Add("其他障礙", "D");

            // 0 一般生
            Dictionary<string, string> g3ItemList = new Dictionary<string, string>();
            g3ItemList.Add("身障生", "1");
            g3ItemList.Add("原住民(有認證)", "2");
            g3ItemList.Add("原住民(無認證)", "3");
            g3ItemList.Add("蒙藏生", "4");
            g3ItemList.Add("外派子女25%", "5");
            g3ItemList.Add("外派子女15%", "6");
            g3ItemList.Add("外派子女10%", "7");
            g3ItemList.Add("退伍軍人25%", "8");
            g3ItemList.Add("退伍軍人20%", "9");
            g3ItemList.Add("退伍軍人15%", "A");
            g3ItemList.Add("退伍軍人10%", "B");
            g3ItemList.Add("退伍軍人5%", "C");
            g3ItemList.Add("退伍軍人3%", "D");
            g3ItemList.Add("優秀子女25%", "E");
            g3ItemList.Add("優秀子女15%", "F");
            g3ItemList.Add("優秀子女10%", "G");
            g3ItemList.Add("僑生", "H");


            foreach (string gname in gList)
            {
                XElement g1 = new XElement("Group");
                g1.SetAttributeValue("Name", gname);

                if (gname == "學生身分")
                {
                    foreach (string na in g1ItemList.Keys)
                    {
                        XElement elm = new XElement("Item");
                        elm.SetAttributeValue("Name", na);
                        elm.SetAttributeValue("TagName", "");
                        elm.SetAttributeValue("Code", g1ItemList[na]);
                        g1.Add(elm);
                    }
                }

                if (gname == "身心障礙")
                {
                    foreach (string na in g2ItemList.Keys)
                    {
                        XElement elm = new XElement("Item");
                        elm.SetAttributeValue("Name", na);
                        elm.SetAttributeValue("TagName", "");
                        elm.SetAttributeValue("Code", g2ItemList[na]);
                        g1.Add(elm);
                    }
                }
                if (gname == "學生報名身分設定")
                {
                    foreach (string na in g3ItemList.Keys)
                    {
                        XElement elm = new XElement("Item");
                        elm.SetAttributeValue("Name", na);
                        elm.SetAttributeValue("TagName", "");
                        elm.SetAttributeValue("Code", g3ItemList[na]);
                        g1.Add(elm);
                    }
                }

                if (gname == "失業勞工子女")
                {
                    XElement elm = new XElement("Item");
                    elm.SetAttributeValue("Name", "是");
                    elm.SetAttributeValue("TagName", "");
                    elm.SetAttributeValue("Code", "1");
                    g1.Add(elm);
                }

                elmRoot.Add(g1);
            }



            return elmRoot;
        }

        private void UserControlEnable(bool value)
        {
            btnSave.Enabled = value;
            btnExport.Enabled = value;
        }
    }
}
