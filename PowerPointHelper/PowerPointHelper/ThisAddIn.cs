using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;

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
            // 이벤트 추가
            this.Application.SlideSelectionChanged -= Application_SlideSelectionChanged;
            this.Application.WindowBeforeRightClick -= Application_WindowBeforeRightClick;
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
            this.Application.WindowBeforeRightClick += Application_WindowBeforeRightClick;

            // 리소스 업데이트
            UpdateResources();
        }

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out System.Drawing.Point lpPoint);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private Form hostFormInstance;
        // ContextMenuStrip 인스턴스를 추적하기 위한 변수
        private ContextMenuStrip hostContextMenuStrip;


        //protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject() {
        //    return new ShapeContextMenu();
        //}

        public void Application_WindowBeforeRightClick(PowerPoint.Selection sel, ref bool Cancel) {
            int VK_SHIFT = 0x10;
            if ((GetAsyncKeyState(VK_SHIFT)) == 0) {
                // Cancel을 true로 바꾸지 않았으므로, 파워포인트의 기본 메뉴가 나타납니다.
                return;
            }

            Cancel = true;
            // 이전에 열려있던 메뉴를 먼저 닫습니다. (안정성을 위함)
            hostContextMenuStrip?.Close();
            if (sel.Type == PpSelectionType.ppSelectionSlides) {


            }

            // 1. 마스터의 UserControl 인스턴스를 생성합니다.
            var myUserControl = new HelperContextMenu();

            // 2. UserControl을 감싸서 메뉴 항목처럼 만들어 줄 ToolStripControlHost를 생성합니다.
            var controlHost = new ToolStripControlHost(myUserControl) {
                // 불필요한 여백을 제거하여 UserControl이 꽉 차게 보입니다.
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                AutoSize = true, // UserControl 크기에 자동으로 맞춰집니다.
                BackColor = Color.Black,
                ForeColor = Color.Black
            };


            // 3. 최종 컨텍스트 메뉴인 ContextMenuStrip을 생성합니다.
            hostContextMenuStrip = new ContextMenuStrip();
            hostContextMenuStrip.BackColor = Color.Black;

            // 4. 메뉴 스트립에 UserControl을 담은 호스트를 추가합니다.
            hostContextMenuStrip.Items.Add(controlHost);

            // 5. 마우스 위치에 ContextMenuStrip을 띄웁니다.
            System.Drawing.Point cursor;
            GetCursorPos(out cursor);
            hostContextMenuStrip.Show(cursor.X, cursor.Y);


            //ContextMenuForm formef = new ContextMenuForm();

            //var myUserControl = new HelperContextMenu();
            //formef.Controls.Add(myUserControl);
            //System.Drawing.Point cursor;
            //GetCursorPos(out cursor);
            //formef.Location = new System.Drawing.Point(cursor.X, cursor.Y);

            //formef.Show();
            //formef.Activate();


            //if (hostFormInstance != null && hostFormInstance.Visible) {
            //    System.Drawing.Point tmepcur;
            //    GetCursorPos(out tmepcur);
            //    hostFormInstance.Location = new System.Drawing.Point(tmepcur.X, tmepcur.Y);
            //    return;
            //}
            //// [재클릭 해결] 이전 메뉴가 있다면 '반드시' 먼저 닫습니다.
            //hostFormInstance?.Close();

            //// [사이즈 해결] 새로운 임시 폼을 생성하되, AutoSize 속성을 사용합니다.
            //hostFormInstance = new Form {
            //    FormBorderStyle = FormBorderStyle.None,
            //    ShowInTaskbar = false,
            //    TopMost = true,
            //    StartPosition = FormStartPosition.Manual,
            //    // 폼이 내용물에 맞게 자동으로 크기를 조절하도록 설정합니다.
            //    AutoSize = true,
            //    AutoSizeMode = AutoSizeMode.GrowAndShrink
            //};

            //ContextMenuForm formef = new ContextMenuForm();
            //formef.Show();

            //hostFormInstance.Deactivate += (s, e) => {
            //    hostFormInstance.Close();
            //};

            //// 새로운 UserControl 인스턴스를 폼에 추가합니다.
            //var myUserControl = new HelperContextMenu();
            //hostFormInstance.Controls.Add(myUserControl);

            //// AutoSize가 켜져 있으므로, 폼은 이제 자동으로 UserControl의 크기에 맞춰집니다.

            //System.Drawing.Point cursor;
            //GetCursorPos(out cursor);
            //hostFormInstance.Location = new System.Drawing.Point(cursor.X, cursor.Y);

            //hostFormInstance.Show();
            //// [재클릭 해결] 폼을 강제로 활성화하여 포커스를 받도록 합니다.
            //hostFormInstance.Activate();
        }

        private void Application_SlideSelectionChanged(PowerPoint.SlideRange SldRange) {
            helperRibbon.Update();

            int VK_SHIFT = 0x10;
            if ((GetAsyncKeyState(VK_SHIFT)) == 0) {
                // Cancel을 true로 바꾸지 않았으므로, 파워포인트의 기본 메뉴가 나타납니다.
                return;
            }

            hostContextMenuStrip?.Close();
            // 1. 마스터의 UserControl 인스턴스를 생성합니다.
            var myUserControl = new HelperContextMenu();

            // 2. UserControl을 감싸서 메뉴 항목처럼 만들어 줄 ToolStripControlHost를 생성합니다.
            var controlHost = new ToolStripControlHost(myUserControl) {
                // 불필요한 여백을 제거하여 UserControl이 꽉 차게 보입니다.
                Margin = Padding.Empty,
                Padding = Padding.Empty,
                AutoSize = true, // UserControl 크기에 자동으로 맞춰집니다.
                BackColor = Color.Black,
                ForeColor = Color.Black
            };


            // 3. 최종 컨텍스트 메뉴인 ContextMenuStrip을 생성합니다.
            hostContextMenuStrip = new ContextMenuStrip();
            hostContextMenuStrip.BackColor = Color.Black;

            // 4. 메뉴 스트립에 UserControl을 담은 호스트를 추가합니다.
            hostContextMenuStrip.Items.Add(controlHost);

            // 5. 마우스 위치에 ContextMenuStrip을 띄웁니다.
            System.Drawing.Point cursor;
            GetCursorPos(out cursor);
            hostContextMenuStrip.Show(cursor.X, cursor.Y);
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
