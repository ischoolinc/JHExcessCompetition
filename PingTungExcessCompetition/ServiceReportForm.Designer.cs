﻿namespace PingTungExcessCompetition
{
    partial class ServiceReportForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lnkViewMapColumns = new System.Windows.Forms.LinkLabel();
            this.lnkViewTemplate = new System.Windows.Forms.LinkLabel();
            this.lnkChangeTemplate = new System.Windows.Forms.LinkLabel();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            this.lnkDefault = new System.Windows.Forms.LinkLabel();
            this.SuspendLayout();
            // 
            // lnkViewMapColumns
            // 
            this.lnkViewMapColumns.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lnkViewMapColumns.AutoSize = true;
            this.lnkViewMapColumns.BackColor = System.Drawing.Color.Transparent;
            this.lnkViewMapColumns.Location = new System.Drawing.Point(202, 43);
            this.lnkViewMapColumns.Name = "lnkViewMapColumns";
            this.lnkViewMapColumns.Size = new System.Drawing.Size(112, 17);
            this.lnkViewMapColumns.TabIndex = 2;
            this.lnkViewMapColumns.TabStop = true;
            this.lnkViewMapColumns.Text = "檢視合併欄位總表";
            this.lnkViewMapColumns.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkViewMapColumns_LinkClicked);
            // 
            // lnkViewTemplate
            // 
            this.lnkViewTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lnkViewTemplate.AutoSize = true;
            this.lnkViewTemplate.BackColor = System.Drawing.Color.Transparent;
            this.lnkViewTemplate.Location = new System.Drawing.Point(23, 43);
            this.lnkViewTemplate.Name = "lnkViewTemplate";
            this.lnkViewTemplate.Size = new System.Drawing.Size(86, 17);
            this.lnkViewTemplate.TabIndex = 0;
            this.lnkViewTemplate.TabStop = true;
            this.lnkViewTemplate.Text = "檢視套印樣板";
            this.lnkViewTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkViewTemplate_LinkClicked);
            // 
            // lnkChangeTemplate
            // 
            this.lnkChangeTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lnkChangeTemplate.AutoSize = true;
            this.lnkChangeTemplate.BackColor = System.Drawing.Color.Transparent;
            this.lnkChangeTemplate.Location = new System.Drawing.Point(112, 43);
            this.lnkChangeTemplate.Name = "lnkChangeTemplate";
            this.lnkChangeTemplate.Size = new System.Drawing.Size(86, 17);
            this.lnkChangeTemplate.TabIndex = 1;
            this.lnkChangeTemplate.TabStop = true;
            this.lnkChangeTemplate.Text = "變更套印樣板";
            this.lnkChangeTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkChangeTemplate_LinkClicked);
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnPrint.Enabled = false;
            this.btnPrint.Location = new System.Drawing.Point(253, 89);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(67, 23);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 4;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(333, 89);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(67, 23);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "離開";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lnkDefault
            // 
            this.lnkDefault.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lnkDefault.AutoSize = true;
            this.lnkDefault.BackColor = System.Drawing.Color.Transparent;
            this.lnkDefault.Location = new System.Drawing.Point(320, 43);
            this.lnkDefault.Name = "lnkDefault";
            this.lnkDefault.Size = new System.Drawing.Size(86, 17);
            this.lnkDefault.TabIndex = 3;
            this.lnkDefault.TabStop = true;
            this.lnkDefault.Text = "檢視預設樣板";
            this.lnkDefault.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkDefault_LinkClicked);
            // 
            // ServiceReportForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(417, 121);
            this.Controls.Add(this.lnkDefault);
            this.Controls.Add(this.lnkViewMapColumns);
            this.Controls.Add(this.lnkViewTemplate);
            this.Controls.Add(this.lnkChangeTemplate);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.btnCancel);
            this.DoubleBuffered = true;
            this.Name = "ServiceReportForm";
            this.Text = "班級服務表現";
            this.Load += new System.EventHandler(this.ServiceReportForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.LinkLabel lnkViewMapColumns;
        private System.Windows.Forms.LinkLabel lnkViewTemplate;
        private System.Windows.Forms.LinkLabel lnkChangeTemplate;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnCancel;
        private System.Windows.Forms.LinkLabel lnkDefault;
    }
}