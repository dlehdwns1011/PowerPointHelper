﻿using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
            this.BookMarkGroup = this.Factory.CreateRibbonGroup();
            this.SetBookMarkButton = this.Factory.CreateRibbonButton();
            this.RemoveBookMarkButton = this.Factory.CreateRibbonButton();
            this.BookMarkListButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.BookMarkGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.BookMarkGroup);
            this.tab1.Label = "Helper";
            this.tab1.Name = "tab1";
            // 
            // BookMarkGroup
            // 
            this.BookMarkGroup.Items.Add(this.SetBookMarkButton);
            this.BookMarkGroup.Items.Add(this.RemoveBookMarkButton);
            this.BookMarkGroup.Items.Add(this.BookMarkListButton);
            this.BookMarkGroup.Label = global::PowerPointHelper.Properties.Resources.RID_BookMark;
            this.BookMarkGroup.Name = "BookMarkGroup";
            // 
            // SetBookMarkButton
            // 
            this.SetBookMarkButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SetBookMarkButton.Image = global::PowerPointHelper.Properties.Resources.BookMarkImage;
            this.SetBookMarkButton.Label = global::PowerPointHelper.Properties.Resources.RID_SetBookMark;
            this.SetBookMarkButton.Name = "SetBookMarkButton";
            this.SetBookMarkButton.ScreenTip = global::PowerPointHelper.Properties.Resources.RID_TipSetBookMark;
            this.SetBookMarkButton.ShowImage = true;
            this.SetBookMarkButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetBookMarkButton_Click);
            // 
            // RemoveBookMarkButton
            // 
            this.RemoveBookMarkButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.RemoveBookMarkButton.Enabled = false;
            this.RemoveBookMarkButton.Image = global::PowerPointHelper.Properties.Resources.BookMarkImage2;
            this.RemoveBookMarkButton.Label = global::PowerPointHelper.Properties.Resources.RID_RemoveBookMark;
            this.RemoveBookMarkButton.Name = "RemoveBookMarkButton";
            this.RemoveBookMarkButton.ScreenTip = global::PowerPointHelper.Properties.Resources.RID_TipRemoveBookMark;
            this.RemoveBookMarkButton.ShowImage = true;
            this.RemoveBookMarkButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RemoveBookMarkButton_Click);
            // 
            // BookMarkListButton
            // 
            this.BookMarkListButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.BookMarkListButton.Image = global::PowerPointHelper.Properties.Resources.BookMarkList;
            this.BookMarkListButton.Label = global::PowerPointHelper.Properties.Resources.RID_BookMarkList;
            this.BookMarkListButton.Name = "BookMarkListButton";
            this.BookMarkListButton.ScreenTip = global::PowerPointHelper.Properties.Resources.RID_TipBookMarkList;
            this.BookMarkListButton.ShowImage = true;
            this.BookMarkListButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BookMarkListButton_Click);
            // 
            // HelperRibbon
            // 
            this.Name = "HelperRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.HelperRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.BookMarkGroup.ResumeLayout(false);
            this.BookMarkGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup BookMarkGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SetBookMarkButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton RemoveBookMarkButton;
        internal RibbonButton BookMarkListButton;
    }

    partial class ThisRibbonCollection {
        internal HelperRibbon HelperRibbon
        {
            get { return this.GetRibbon<HelperRibbon>(); }
        }
    }
}
