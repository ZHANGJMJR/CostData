using System;

namespace CostData
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            initDateControl();
        }

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
        public DateTime[] initDate()
        {
            return new DateTime[2] { DateTime.Now.AddDays(-1), DateTime.Now.AddDays(-1) };
        }
        public string initLastMonth()
        {
            DateTime currentDate = DateTime.Now;
            DateTime lastMonth = currentDate.AddMonths(-1);
            return lastMonth.ToString("yyyy-MM");
        }

        public void initDateControl()
        {
            this.editBox1.Text = initDate()[0].ToString("yyyy-MM-dd");
            this.editBox2.Text = initDate()[1].ToString("yyyy-MM-dd");
            this.banquetdate.Text = initLastMonth();
            this.chinesefoodeditBox.Text = initLastMonth();
        }
        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.editBox1 = this.Factory.CreateRibbonEditBox();
            this.box1 = this.Factory.CreateRibbonBox();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.editBox2 = this.Factory.CreateRibbonEditBox();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.download_btn = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.banquetdate = this.Factory.CreateRibbonEditBox();
            this.banquetbtn = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.chinesefoodeditBox = this.Factory.CreateRibbonEditBox();
            this.chinesefoodbtn = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "审计数据处理";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.editBox1);
            this.group1.Items.Add(this.box1);
            this.group1.Items.Add(this.editBox2);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.download_btn);
            this.group1.Name = "group1";
            // 
            // editBox1
            // 
            this.editBox1.Label = "开始日期";
            this.editBox1.MaxLength = 20;
            this.editBox1.Name = "editBox1";
            this.editBox1.Text = null;
            this.editBox1.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox1_TextChanged);
            // 
            // box1
            // 
            this.box1.Items.Add(this.label1);
            this.box1.Name = "box1";
            // 
            // label1
            // 
            this.label1.Label = "|";
            this.label1.Name = "label1";
            // 
            // editBox2
            // 
            this.editBox2.Label = "结束日期";
            this.editBox2.MaxLength = 20;
            this.editBox2.Name = "editBox2";
            this.editBox2.Text = null;
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // download_btn
            // 
            this.download_btn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.download_btn.Image = global::CostData.Properties.Resources.checkbill;
            this.download_btn.Label = "双系统账单对账";
            this.download_btn.Name = "download_btn";
            this.download_btn.ShowImage = true;
            this.download_btn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.download_btn_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.banquetdate);
            this.group3.Items.Add(this.banquetbtn);
            this.group3.Name = "group3";
            // 
            // banquetdate
            // 
            this.banquetdate.Label = "月份";
            this.banquetdate.MaxLength = 7;
            this.banquetdate.Name = "banquetdate";
            this.banquetdate.ScreenTip = "输入月份";
            this.banquetdate.SuperTip = "输入月份";
            this.banquetdate.Text = null;
            // 
            // banquetbtn
            // 
            this.banquetbtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.banquetbtn.Image = global::CostData.Properties.Resources.banquet;
            this.banquetbtn.Label = "宴会数据";
            this.banquetbtn.Name = "banquetbtn";
            this.banquetbtn.ShowImage = true;
            this.banquetbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.banquetbtn_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.chinesefoodeditBox);
            this.group4.Items.Add(this.chinesefoodbtn);
            this.group4.Name = "group4";
            // 
            // chinesefoodeditBox
            // 
            this.chinesefoodeditBox.Label = "月份";
            this.chinesefoodeditBox.MaxLength = 7;
            this.chinesefoodeditBox.Name = "chinesefoodeditBox";
            this.chinesefoodeditBox.ScreenTip = "输入月份";
            this.chinesefoodeditBox.SuperTip = "输入月份";
            this.chinesefoodeditBox.Text = null;
            // 
            // chinesefoodbtn
            // 
            this.chinesefoodbtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.chinesefoodbtn.Image = global::CostData.Properties.Resources.chinesefood;
            this.chinesefoodbtn.Label = "中餐数据";
            this.chinesefoodbtn.Name = "chinesefoodbtn";
            this.chinesefoodbtn.ShowImage = true;
            this.chinesefoodbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chinesefoodbtn_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button2);
            this.group2.Label = "关于";
            this.group2.Name = "group2";
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::CostData.Properties.Resources.logo;
            this.button2.Label = " ";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }
        
        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton download_btn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox1;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox editBox2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton banquetbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox banquetdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox chinesefoodeditBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton chinesefoodbtn;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
