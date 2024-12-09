using System;
using System.Collections.Generic;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointHelper {
    public partial class BookMarkListDlg : Form {
        public BookMarkListDlg() {
            InitializeComponent();
            Init();
        }

        List<int> sldList;

        private void Init() {
            this.bookMarkListBox.Items.Clear();
            sldList = Globals.ThisAddIn.bookMarkManager.GetBookMarkedSlideIndex();

            var slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            foreach (int sld in sldList) {
                this.bookMarkListBox.Items.Add("[" + sld.ToString() + "] " + slides[sld].Tags["bookmark"]);
            }

            if (this.bookMarkListBox.SelectedItems.Count == 0) {
                this.editBookMark.Enabled = false;
                this.removeBookMark.Enabled = false;
                this.goToBookMark.Enabled = false;
            } else if (this.bookMarkListBox.SelectedItems.Count == 1) {
                this.editBookMark.Enabled = true;
                this.removeBookMark.Enabled = true;
                this.goToBookMark.Enabled = true;
            } else if (this.bookMarkListBox.SelectedItems.Count > 1) {
                this.editBookMark.Enabled = false;
                this.removeBookMark.Enabled = true;
                this.goToBookMark.Enabled = false;
            }
        }

        private void editBookMark_Click(object sender, EventArgs e) {
            var listBox = this.bookMarkListBox;
            if (listBox == null) {
                return;
            }

            SetBookMarkDlg setBookMarkDlg = new SetBookMarkDlg();
            setBookMarkDlg.Text = Properties.Resources.RID_BookMarkEdit;
            setBookMarkDlg.selectedSlideIndex = sldList[listBox.SelectedIndex];
            setBookMarkDlg.SetBeforeName(Globals.ThisAddIn.Application.ActivePresentation.Slides[sldList[listBox.SelectedIndex]].Tags["bookmark"]);
            setBookMarkDlg.ShowDialog();

            if (setBookMarkDlg.DialogResult == DialogResult.OK) {
                Init();
            }
        }

        private void removeBookMark_Click(object sender, EventArgs e) {
            var listBox = this.bookMarkListBox;
            if (listBox == null) {
                return;
            }

            List<int> removeList = new List<int>();
            for (int index = 0; index < listBox.SelectedIndices.Count;++index) {
                removeList.Add(sldList[listBox.SelectedIndices[index]]);
            }

            Globals.ThisAddIn.bookMarkManager.DeleteBookMarks(removeList);

            Init();
        }

        private void goToBookMark_Click(object sender, EventArgs e) {
            var listBox = this.bookMarkListBox;
            if (listBox == null) {
                return;
            }

            this.Close();
            Globals.ThisAddIn.bookMarkManager.MoveBookMark(sldList[listBox.SelectedIndex]);
        }

        private void bookMarkListBox_SelectedIndexChanged(object sender, EventArgs e) {
            // 책갈피 목록 리스트 변경 이벤트
            var listBox = sender as ListBox;
            if (listBox == null) {
                return;
            }

            if (listBox.SelectedItems.Count == 0) {
                this.editBookMark.Enabled = false;
                this.removeBookMark.Enabled = false;
                this.goToBookMark.Enabled = false;
            } else if (listBox.SelectedItems.Count == 1) {
                this.editBookMark.Enabled = true;
                this.removeBookMark.Enabled = true;
                this.goToBookMark.Enabled = true;
            } else if (listBox.SelectedItems.Count > 1) {
                this.editBookMark.Enabled = false;
                this.removeBookMark.Enabled = true;
                this.goToBookMark.Enabled = false;
            }
        }

        private void bookMarkListBox_DoulbeClick(object sender, MouseEventArgs e) {
            var listBox = sender as ListBox;
            if (listBox == null || listBox.SelectedItems.Count != 1) {
                return;
            }

            var selectedItemRect = listBox.GetItemRectangle(listBox.SelectedIndex);
            if(selectedItemRect.Contains(e.X, e.Y)) {
                this.Close();
                Globals.ThisAddIn.bookMarkManager.MoveBookMark(sldList[listBox.SelectedIndex]);
            }

        }
    }
}
