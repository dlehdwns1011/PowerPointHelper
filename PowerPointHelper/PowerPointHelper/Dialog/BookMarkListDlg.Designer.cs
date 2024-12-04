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
            this.bookMarkListBox = new System.Windows.Forms.ListBox();
            this.editBookMark = new System.Windows.Forms.Button();
            this.removeBookMark = new System.Windows.Forms.Button();
            this.goToBookMark = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // bookMarkListBox
            // 
            this.bookMarkListBox.FormattingEnabled = true;
            this.bookMarkListBox.ItemHeight = 12;
            this.bookMarkListBox.Location = new System.Drawing.Point(13, 13);
            this.bookMarkListBox.Name = "bookMarkListBox";
            this.bookMarkListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.bookMarkListBox.Size = new System.Drawing.Size(292, 304);
            this.bookMarkListBox.TabIndex = 0;
            this.bookMarkListBox.SelectedIndexChanged += new System.EventHandler(this.bookMarkListBox_SelectedIndexChanged);
            this.bookMarkListBox.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.bookMarkListBox_DoulbeClick);
            // 
            // editBookMark
            // 
            this.editBookMark.Enabled = false;
            this.editBookMark.Location = new System.Drawing.Point(15, 333);
            this.editBookMark.Name = "editBookMark";
            this.editBookMark.Size = new System.Drawing.Size(67, 30);
            this.editBookMark.TabIndex = 1;
            this.editBookMark.Text = global::PowerPointHelper.Properties.Resources.RID_Edit;
            this.editBookMark.UseVisualStyleBackColor = true;
            this.editBookMark.Click += new System.EventHandler(this.editBookMark_Click);
            // 
            // removeBookMark
            // 
            this.removeBookMark.Enabled = false;
            this.removeBookMark.Location = new System.Drawing.Point(88, 333);
            this.removeBookMark.Name = "removeBookMark";
            this.removeBookMark.Size = new System.Drawing.Size(67, 30);
            this.removeBookMark.TabIndex = 2;
            this.removeBookMark.Text = global::PowerPointHelper.Properties.Resources.RID_Remove;
            this.removeBookMark.UseVisualStyleBackColor = true;
            this.removeBookMark.Click += new System.EventHandler(this.removeBookMark_Click);
            // 
            // goToBookMark
            // 
            this.goToBookMark.Enabled = false;
            this.goToBookMark.Location = new System.Drawing.Point(161, 333);
            this.goToBookMark.Name = "goToBookMark";
            this.goToBookMark.Size = new System.Drawing.Size(67, 30);
            this.goToBookMark.TabIndex = 3;
            this.goToBookMark.Text = global::PowerPointHelper.Properties.Resources.RID_Move;
            this.goToBookMark.UseVisualStyleBackColor = true;
            this.goToBookMark.Click += new System.EventHandler(this.goToBookMark_Click);
            // 
            // BookMarkListDlg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(319, 380);
            this.Controls.Add(this.goToBookMark);
            this.Controls.Add(this.removeBookMark);
            this.Controls.Add(this.editBookMark);
            this.Controls.Add(this.bookMarkListBox);
            this.Name = "BookMarkListDlg";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = global::PowerPointHelper.Properties.Resources.RID_BookMarkList;
            this.ResumeLayout(false);
            this.ShowIcon = false;

        }

        #endregion

        private System.Windows.Forms.ListBox bookMarkListBox;
        
        private System.Windows.Forms.Button editBookMark;
        private System.Windows.Forms.Button removeBookMark;
        private System.Windows.Forms.Button goToBookMark;
    }
}