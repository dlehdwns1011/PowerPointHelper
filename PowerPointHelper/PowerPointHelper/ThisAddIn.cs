using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointHelper
{
    public partial class ThisAddIn
    {
        public HelperRibbon helperRibbon { get; set; }
        public BookMarkManager bookMarkManager { get; set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            init();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region -> private 함수
        private void init() {
            // 언어 설정
            System.Globalization.CultureInfo cultureInfo = 
                new System.Globalization.CultureInfo(this.Application.LanguageSettings.LanguageID[Office.MsoAppLanguageID.msoLanguageIDUI]);
            System.Threading.Thread.CurrentThread.CurrentCulture = cultureInfo;
            System.Threading.Thread.CurrentThread.CurrentUICulture = cultureInfo;

            helperRibbon = Globals.Ribbons.HelperRibbon;
            bookMarkManager = new BookMarkManager();

            // 이벤트 추가
            this.Application.SlideSelectionChanged += Application_SlideSelectionChanged;

            // 리소스 업데이트
            UpdateResources();
        }

        private void Application_SlideSelectionChanged(PowerPoint.SlideRange SldRange) {
            helperRibbon.Update();
        }

        private void UpdateResources() {
            this.helperRibbon.UpdateResources();
        }
        #endregion

        #region VSTO에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
