namespace Daisy.SaveAsDAISY.Addins.Word2007
{
    partial class Abbreviation
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.lBx_Abbreviation = new System.Windows.Forms.ListBox();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.btn_Unmark = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btn_Goto = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lBx_Abbreviation
            // 
            this.lBx_Abbreviation.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lBx_Abbreviation.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lBx_Abbreviation.FormattingEnabled = true;
            this.lBx_Abbreviation.HorizontalScrollbar = true;
            this.lBx_Abbreviation.ItemHeight = 15;
            this.lBx_Abbreviation.Location = new System.Drawing.Point(6, 25);
            this.lBx_Abbreviation.Name = "lBx_Abbreviation";
            this.lBx_Abbreviation.Size = new System.Drawing.Size(256, 167);
            this.lBx_Abbreviation.TabIndex = 1;
            this.lBx_Abbreviation.SelectedValueChanged += new System.EventHandler(this.lBx_Abbreviation_SelectedValueChanged);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btn_Cancel.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btn_Cancel.Location = new System.Drawing.Point(187, 199);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(75, 23);
            this.btn_Cancel.TabIndex = 3;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = false;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // btn_Unmark
            // 
            this.btn_Unmark.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btn_Unmark.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btn_Unmark.Location = new System.Drawing.Point(108, 199);
            this.btn_Unmark.Name = "btn_Unmark";
            this.btn_Unmark.Size = new System.Drawing.Size(75, 23);
            this.btn_Unmark.TabIndex = 2;
            this.btn_Unmark.Text = "&Unmark";
            this.btn_Unmark.UseVisualStyleBackColor = false;
            this.btn_Unmark.Click += new System.EventHandler(this.btn_Unmark_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.label1.Location = new System.Drawing.Point(3, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 15);
            this.label1.TabIndex = 1;
            // 
            // btn_Goto
            // 
            this.btn_Goto.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btn_Goto.Location = new System.Drawing.Point(27, 199);
            this.btn_Goto.Name = "btn_Goto";
            this.btn_Goto.Size = new System.Drawing.Size(75, 23);
            this.btn_Goto.TabIndex = 5;
            this.btn_Goto.Text = "&GoTo";
            this.btn_Goto.UseVisualStyleBackColor = false;
            this.btn_Goto.Visible = false;
            this.btn_Goto.Click += new System.EventHandler(this.btn_Goto_Click);
            // 
            // Abbreviation
            // 
            this.AcceptButton = this.btn_Unmark;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.CancelButton = this.btn_Cancel;
            this.ClientSize = new System.Drawing.Size(272, 225);
            this.Controls.Add(this.btn_Goto);
            this.Controls.Add(this.lBx_Abbreviation);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_Unmark);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Abbreviation";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Abbreviation_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lBx_Abbreviation;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Button btn_Unmark;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.Button btn_Goto;
    }
}