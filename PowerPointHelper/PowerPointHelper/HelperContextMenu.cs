using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointHelper {
    public partial class HelperContextMenu : UserControl {
        public HelperContextMenu() {
            InitializeComponent();

            this.Padding = new Padding(0);
            this.Margin = new Padding(0);
        }

        private void button1_Click(object sender, EventArgs e) {
            MessageBox.Show("hello1");

        }

        private bool Button1_IsEnagleAddBookMark() {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.Count > 1) {
                return false;
            }

            var nowSlide = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange[1];

            if (Globals.ThisAddIn.bookMarkManager.IsExistBookMark(nowSlide)) {
                return false;
            }

            return true;
        }

        private bool Button1_IsEnagleDeleteBookMark() {
            if (Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange.Count > 1) {
                return false;
            }

            var nowSlide = Globals.ThisAddIn.Application.ActiveWindow.Selection.SlideRange[1];

            if (Globals.ThisAddIn.bookMarkManager.IsExistBookMark(nowSlide)) {
                return true;
            }

            return false;
        }
    }
}
