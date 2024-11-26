namespace PowerPointHelper {
    partial class HelperRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public HelperRibbon()
            : base(Globals.Factory.GetRibbonFactory()) {
            InitializeComponent();
        }

        /// <summary> 
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 구성 요소 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent() {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.SetBookMarkButton = this.Factory.CreateRibbonButton();
            this.RemoveBookMarkButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "Helper";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.SetBookMarkButton);
            this.group1.Items.Add(this.RemoveBookMarkButton);
            this.group1.Label = "책갈피";
            this.group1.Name = "group1";
            // 
            // SetBookMarkButton
            // 
            this.SetBookMarkButton.Image = global::PowerPointHelper.Properties.Resources.BookMarkImage;
            this.SetBookMarkButton.Label = "책갈피 추가";
            this.SetBookMarkButton.Name = "SetBookMarkButton";
            this.SetBookMarkButton.ShowImage = true;
            this.SetBookMarkButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            // 
            // RemoveBookMarkButton
            // 
            this.RemoveBookMarkButton.Image = global::PowerPointHelper.Properties.Resources.BookMarkImage;
            this.RemoveBookMarkButton.Label = "책갈피 제거";
            this.RemoveBookMarkButton.Name = "RemoveBookMarkButton";
            this.RemoveBookMarkButton.ShowImage = true;
            this.RemoveBookMarkButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            // 
            // HelperRibbon
            // 
            this.Name = "HelperRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.HelperRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SetBookMarkButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RemoveBookMarkButton;
    }

    partial class ThisRibbonCollection {
        internal HelperRibbon HelperRibbon
        {
            get { return this.GetRibbon<HelperRibbon>(); }
        }
    }
}
