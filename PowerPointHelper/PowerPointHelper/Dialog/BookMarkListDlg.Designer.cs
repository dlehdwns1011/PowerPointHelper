namespace PowerPointHelper {
    partial class BookMarkListDlg {
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
            this.bookMarkListBox = new System.Windows.Forms.CheckedListBox();
            this.SuspendLayout();
            // 
            // bookMarkListBox
            // 
            this.bookMarkListBox.FormattingEnabled = true;
            this.bookMarkListBox.Location = new System.Drawing.Point(13, 13);
            this.bookMarkListBox.Name = "bookMarkListBox";
            this.bookMarkListBox.Size = new System.Drawing.Size(292, 308);
            this.bookMarkListBox.TabIndex = 0;
            // 
            // BookMarkListDlg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.bookMarkListBox);
            this.Name = "BookMarkListDlg";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "책갈피 목록";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckedListBox bookMarkListBox;
    }
}