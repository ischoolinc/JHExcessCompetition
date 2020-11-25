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
using FISCA.UDT;
using Aspose.Words;
using System.IO;
using System.Xml.Linq;

namespace PingTungExcessCompetition
{
    public partial class setCadreNameForm : BaseForm
    {
        public Configure _Configure { get; private set; }
        AccessHelper _accessHelper = new AccessHelper();

        public setCadreNameForm()
        {
            InitializeComponent();
            btnSave.DialogResult = DialogResult.OK;
            btnExit.DialogResult = DialogResult.Cancel;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveData();           
        }

        private void SaveData()
        {
            try
            {
                List<string> nameList = new List<string>();
                foreach (DataGridViewRow drv in dgData.Rows)
                {
                    if (drv.IsNewRow)
                        continue;

                    string name = drv.Cells[colName.Index].Value + "";

                    if (!nameList.Contains(name))
                        nameList.Add(name);
                }

                XElement elmRoot = new XElement("Items");
                foreach(string name in nameList)
                {
                    XElement elm = new XElement("Item");
                    elm.SetAttributeValue("Name", name);
                    elmRoot.Add(elm);
                }
                _Configure.CadreNames = elmRoot.ToString();
                _Configure.Save();
                MessageBox.Show("儲存完成");
                this.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show("儲存失敗:" + ex.Message);
            }

        }

        private void setCadreNameForm_Load(object sender, EventArgs e)
        {
            btnSave.Enabled = false;
            LoadDataToForm();

        }

        private void LoadDataToForm()
        {
            dgData.Rows.Clear();
            List<string> cDataList = LoadCadreName();
            foreach (string name in cDataList)
            {
                int rowIdx = dgData.Rows.Add();
                dgData.Rows[rowIdx].Cells[colName.Index].Value = name;
            }
            btnSave.Enabled = true;
        }



        private List<string> LoadCadreName()
        {
            List<string> CadreNameList = new List<string>();
            try
            {

                List<Configure> confList = _accessHelper.Select<Configure>();
                if (confList != null && confList.Count > 0)
                {
                    _Configure = confList[0];
                  
                }
                else
                {
                    _Configure = new Configure();
                    _Configure.Name = "屏東免試入學-班級服務表現";
                    _Configure.Template = new Document(new MemoryStream(Properties.Resources.屏東班級服務表現樣板));
                    _Configure.Encode();

                }
                CadreNameList = _Configure.LoadCareNames();
                _Configure.Save();
            }
            catch (Exception ex)
            {
                MsgBox.Show("讀取幹部限制設定發生錯誤，" + ex.Message);
            }
            return CadreNameList;
        }
    }
}
