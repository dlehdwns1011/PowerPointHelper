using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
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

        public void EditBookMark(int sldIndex, string newName) {
            var nowSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides[sldIndex];
            nowSlide.Tags.Delete("bookmark");
            nowSlide.Tags.Add("bookmark", newName);
        }

        public void DeleteBookMarks(List<int> sldList) {
            foreach (int index in sldList) {
                var nowSlide = Globals.ThisAddIn.Application.ActivePresentation.Slides[index];
                nowSlide.Tags.Delete("bookmark");
            }
        }

        public void MoveBookMark(int sldIndex) {
            Globals.ThisAddIn.Application.ActivePresentation.Slides[sldIndex].Select();
        }
    }
}
