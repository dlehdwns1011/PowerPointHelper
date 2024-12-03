using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointHelper {
    partial class SetBookMarkDlg {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.label1 = new System.Windows.Forms.Label();
            this.bookMarkNameText = new System.Windows.Forms.TextBox();
            this.cancelButton = new System.Windows.Forms.Button();
            this.OKButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(81, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "책갈피 이름 : ";
            // 
            // bookMarkNameText
            // 
            this.bookMarkNameText.Location = new System.Drawing.Point(14, 25);
            this.bookMarkNameText.Name = "bookMarkNameText";
            this.bookMarkNameText.Size = new System.Drawing.Size(242, 21);
            this.bookMarkNameText.TabIndex = 1;
            this.bookMarkNameText.TextChanged += new System.EventHandler(this.bookMarkNameText_TextChanged);
            this.bookMarkNameText.KeyUp += bookMarkNameText_KeyUp;
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(181, 52);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(75, 23);
            this.cancelButton.TabIndex = 2;
            this.cancelButton.Text = "취소";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
            // 
            // OKButton
            // 
            this.OKButton.Location = new System.Drawing.Point(100, 52);
            this.OKButton.Name = "OKButton";
            this.OKButton.Size = new System.Drawing.Size(75, 23);
            this.OKButton.TabIndex = 3;
            this.OKButton.Text = "확인";
            this.OKButton.UseVisualStyleBackColor = true;
            this.OKButton.Enabled = false;
            this.OKButton.Click += new System.EventHandler(this.OKButton_Click);
            // 
            // SetBookMarkDlg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(265, 87);
            this.Controls.Add(this.OKButton);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.bookMarkNameText);
            this.Controls.Add(this.label1);
            this.Name = "SetBookMarkDlg";
            this.Text = "책갈피 추가";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.ResumeLayout(false);
            this.PerformLayout();
            this.Icon = new System.Drawing.Icon(Properties.Resources.ResourceManager.GetStream("BookMarkImage"));

        }

        

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox bookMarkNameText;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button OKButton;
    }
}