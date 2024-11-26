using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointHelper {
    public partial class HelperRibbon {
        private void HelperRibbon_Load(object sender, RibbonUIEventArgs e) {

        }

        public void UpdateResources() {
            this.SetBookMarkButton.Label = Properties.Resources.RID_SetBookMark;
            this.RemoveBookMarkButton.Label = Properties.Resources.RID_RemoveBookMark;
            this.BookMarkGroup.Label = Properties.Resources.RID_BookMark;
        }
    }
}
