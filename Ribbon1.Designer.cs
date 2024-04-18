namespace ASRTookit_M
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.reverse = this.Factory.CreateRibbonButton();
            this.Font = this.Factory.CreateRibbonButton();
            this.rm_bg = this.Factory.CreateRibbonButton();
            this.noEmpty = this.Factory.CreateRibbonButton();
            this.error20 = this.Factory.CreateRibbonButton();
            this.error2empty = this.Factory.CreateRibbonButton();
            this.strongnoempty = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Label = "ASRTookit";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Label = "数字格式";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.reverse);
            this.group2.Items.Add(this.Font);
            this.group2.Items.Add(this.rm_bg);
            this.group2.Items.Add(this.noEmpty);
            this.group2.Items.Add(this.error20);
            this.group2.Items.Add(this.error2empty);
            this.group2.Label = "常用数据操作";
            this.group2.Name = "group2";
            // 
            // group3
            // 
            this.group3.Items.Add(this.strongnoempty);
            this.group3.Label = "不常用数据操作";
            this.group3.Name = "group3";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::ASRTookit_M.Properties.Resources.format_red_1;
            this.button1.Label = "格式1";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = global::ASRTookit_M.Properties.Resources.format_red_2;
            this.button2.Label = "格式2";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click_1);
            // 
            // reverse
            // 
            this.reverse.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.reverse.Image = global::ASRTookit_M.Properties.Resources.正负号;
            this.reverse.Label = "负号";
            this.reverse.Name = "reverse";
            this.reverse.ShowImage = true;
            this.reverse.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // Font
            // 
            this.Font.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Font.Image = global::ASRTookit_M.Properties.Resources.YaHei;
            this.Font.Label = "YaHei";
            this.Font.Name = "Font";
            this.Font.ShowImage = true;
            this.Font.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Font_Click);
            // 
            // rm_bg
            // 
            this.rm_bg.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.rm_bg.Image = global::ASRTookit_M.Properties.Resources.清除格式;
            this.rm_bg.Label = "去除色框";
            this.rm_bg.Name = "rm_bg";
            this.rm_bg.ShowImage = true;
            this.rm_bg.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rm_bg_Click);
            // 
            // noEmpty
            // 
            this.noEmpty.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.noEmpty.Image = ((System.Drawing.Image)(resources.GetObject("noEmpty.Image")));
            this.noEmpty.Label = "去空";
            this.noEmpty.Name = "noEmpty";
            this.noEmpty.ShowImage = true;
            this.noEmpty.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // error20
            // 
            this.error20.Label = "错误转0";
            this.error20.Name = "error20";
            this.error20.ScreenTip = "(谨慎全选)公式添加IFERROR为0";
            this.error20.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.error20_Click);
            // 
            // error2empty
            // 
            this.error2empty.Label = "错误转空";
            this.error2empty.Name = "error2empty";
            this.error2empty.ScreenTip = "(谨慎全选)公式添加IFERROR为空";
            this.error2empty.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.error2empty_Click);
            // 
            // strongnoempty
            // 
            this.strongnoempty.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.strongnoempty.Image = ((System.Drawing.Image)(resources.GetObject("strongnoempty.Image")));
            this.strongnoempty.Label = "强力去空";
            this.strongnoempty.Name = "strongnoempty";
            this.strongnoempty.ScreenTip = "(谨慎全选)去除单元格所有空白";
            this.strongnoempty.ShowImage = true;
            this.strongnoempty.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.strongnoempty_Click);
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
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton reverse;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Font;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton noEmpty;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton strongnoempty;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton error20;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton error2empty;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rm_bg;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
