using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointHelper {
    public partial class SetBookMarkDlg : Form {
        public SetBookMarkDlg() {
            InitializeComponent();
            UpdateResources();
        }

        public int selectedSlideIndex { get; set; }

        #region -> 공개 함수
        public void UpdateResources() {
            this.label1.Text = Properties.Resources.RID_BookMarkName;
            this.cancelButton.Text = Properties.Resources.RID_Cancel;
            this.OKButton.Text = Properties.Resources.RID_OK;

            this.Text = Properties.Resources.RID_SetBookMark;
        }

        public void SetBeforeName(String beforeName) {
            this.bookMarkNameText.Text = beforeName;
        }
        #endregion

        private void cancelButton_Click(object sender, EventArgs e) {
            this.Close();
        }

        private void OKButton_Click(object sender, EventArgs e) {
            if (this.Text == Properties.Resources.RID_SetBookMark) {
                // 책갈피 추가하자
                PowerPoint.Slide activeSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                activeSlide.Tags.Add("bookmark", this.bookMarkNameText.Text);

            } else if (this.Text == Properties.Resources.RID_BookMarkEdit) {
                // 책갈피 편집
                this.DialogResult = DialogResult.OK;
                Globals.ThisAddIn.bookMarkManager.EditBookMark(selectedSlideIndex, this.bookMarkNameText.Text);
            }
            
            this.Close();
        }

        private void bookMarkNameText_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter) {
                if (this.bookMarkNameText.Text.Length > 0) {
                    if (this.Text == Properties.Resources.RID_SetBookMark) {
                        // 책갈피 추가하자
                        PowerPoint.Slide activeSlide = Globals.ThisAddIn.Application.ActiveWindow.View.Slide;

                        activeSlide.Tags.Add("bookmark", this.bookMarkNameText.Text);
                    } else if (this.Text == Properties.Resources.RID_BookMarkEdit) {
                        // 책갈피 편집
                        this.DialogResult = DialogResult.OK;
                        Globals.ThisAddIn.bookMarkManager.EditBookMark(selectedSlideIndex, this.bookMarkNameText.Text);
                    }

                    this.Close();
                }
            }
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
