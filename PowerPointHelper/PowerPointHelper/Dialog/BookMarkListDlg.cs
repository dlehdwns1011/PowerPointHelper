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
        }
    }
}
