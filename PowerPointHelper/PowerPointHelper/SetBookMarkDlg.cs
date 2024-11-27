using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointHelper {
    public partial class SetBookMarkDlg : Form {
        public SetBookMarkDlg() {
            InitializeComponent();
            UpdateResources();
        }

        #region -> 공개 함수
        public void UpdateResources() {
            this.label1.Text = Properties.Resources.RID_BookMarkName;
            this.cancelButton.Text = Properties.Resources.RID_Cancel;
            this.OKButton.Text = "확인";

            this.Text = Properties.Resources.RID_SetBookMark;
            
        }
        #endregion

        private void cancelButton_Click(object sender, EventArgs e) {
            this.Close();
        }

        private void OKButton_Click(object sender, EventArgs e) {
            // 책갈피 추가하자
            var activePresentation = Globals.ThisAddIn.Application.ActivePresentation;
            PowerPoint.Slide activeSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            activeSlide.Tags.Add("bookmark", "이거슨 마크");

            this.Close();
        }

        private void bookMarkNameText_TextChanged(object sender, EventArgs e) {
            TextBox textBox = sender as TextBox;
            if (textBox == null) {
                return;
            }

            if (textBox.Text.Length > 0) {
                this.OKButton.Enabled = true;
            } else {
                this.OKButton.Enabled = false;
            }
        }
    }
}
