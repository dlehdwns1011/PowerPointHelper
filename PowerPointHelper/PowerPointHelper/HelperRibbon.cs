using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointHelper {
    public partial class HelperRibbon {
        private void HelperRibbon_Load(object sender, RibbonUIEventArgs e) {

        }

        public void UpdateResources() {
            this.SetBookMarkButton.Label = Properties.Resources.RID_SetBookMark;
            this.RemoveBookMarkButton.Label = Properties.Resources.RID_RemoveBookMark;
            this.BookMarkGroup.Label = Properties.Resources.RID_BookMark;
            this.BookMarkListButton.Label = Properties.Resources.RID_BookMarkList;
        }

        public void Update() {
            PowerPoint.Slide activeSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            if (activeSlide.Tags.Count > 0) {
                this.SetBookMarkButton.Enabled = false;
                this.RemoveBookMarkButton.Enabled = true;
            } else {
                this.SetBookMarkButton.Enabled = true;
                this.RemoveBookMarkButton.Enabled = false;
            }

        }

        private void SetBookMarkButton_Click(object sender, RibbonControlEventArgs e) {
            // 현재 열려있는 슬라이드를 책갈피에 추가
            SetBookMarkDlg dlg = new SetBookMarkDlg();
            dlg.ShowDialog();

            Update();
        }

        private void RemoveBookMarkButton_Click(object sender, RibbonControlEventArgs e) {
            PowerPoint.Slide activeSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            activeSlide.Tags.Delete("bookmark");

            Update();
        }

        private void BookMarkListButton_Click(object sender, RibbonControlEventArgs e) {

        }
    }
}
