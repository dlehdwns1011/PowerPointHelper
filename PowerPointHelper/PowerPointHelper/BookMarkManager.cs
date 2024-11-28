using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointHelper {
    public class BookMarkManager {

        public BookMarkManager() { 

        }

        public List<int> GetBookMarkedSlideIndex() {
            List<int> bookMarkedSlideIndex = new List<int>();
            foreach(Slide sld in Globals.ThisAddIn.Application.ActivePresentation.Slides) {
                if (sld.Tags["bookmark"] != "") {
                    bookMarkedSlideIndex.Add(sld.SlideIndex);
                }
            }

            return bookMarkedSlideIndex;
        }

        public void DeleteBookMarks(List<int> sldList) {
            Globals.ThisAddIn.Application.ActivePresentation.Slides.Range(sldList).Delete();
        }

        public void MoveBookMark(int sldIndex) {
            Globals.ThisAddIn.Application.ActivePresentation.Slides[sldIndex].Select();
        }
    }
}
