namespace Daisy.SaveAsDAISY.Conversion
{
    partial class Fidility
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Fidility));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lbl_Summary = new System.Windows.Forms.Label();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.btn_OK = new System.Windows.Forms.Button();
            this.dataGrid_Fidility = new System.Windows.Forms.DataGridView();
            this.lLbl_Log = new System.Windows.Forms.LinkLabel();
            this.lbl_Information = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid_Fidility)).BeginInit();
            this.SuspendLayout();
            // 
            // lbl_Summary
            // 
            resources.ApplyResources(this.lbl_Summary, "lbl_Summary");
            this.lbl_Summary.Name = "lbl_Summary";
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.BackColor = System.Drawing.SystemColors.ControlLight;
            this.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.btn_Cancel, "btn_Cancel");
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = false;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // btn_OK
            // 
            this.btn_OK.BackColor = System.Drawing.SystemColors.ControlLight;
            resources.ApplyResources(this.btn_OK, "btn_OK");
            this.btn_OK.Name = "btn_OK";
            this.btn_OK.UseVisualStyleBackColor = false;
            this.btn_OK.Click += new System.EventHandler(this.btn_OK_Click);
            // 
            // dataGrid_Fidility
            // 
            this.dataGrid_Fidility.AllowUserToAddRows = false;
            this.dataGrid_Fidility.AllowUserToDeleteRows = false;
            this.dataGrid_Fidility.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGrid_Fidility.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None;
            this.dataGrid_Fidility.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGrid_Fidility.ColumnHeadersVisible = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlLightLight;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGrid_Fidility.DefaultCellStyle = dataGridViewCellStyle1;
            resources.ApplyResources(this.dataGrid_Fidility, "dataGrid_Fidility");
            this.dataGrid_Fidility.Name = "dataGrid_Fidility";
            this.dataGrid_Fidility.ReadOnly = true;
            this.dataGrid_Fidility.RowHeadersVisible = false;
            this.dataGrid_Fidility.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.White;
            this.dataGrid_Fidility.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.dataGrid_Fidility.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGrid_Fidility.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.dataGrid_Fidility_KeyPress);
            this.dataGrid_Fidility.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // lLbl_Log
            // 
            resources.ApplyResources(this.lLbl_Log, "lLbl_Log");
            this.lLbl_Log.LinkColor = System.Drawing.SystemColors.ControlText;
            this.lLbl_Log.Name = "lLbl_Log";
            this.lLbl_Log.TabStop = true;
            this.lLbl_Log.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lLbl_Log_LinkClicked);
            // 
            // lbl_Information
            // 
            this.lbl_Information.BackColor = System.Drawing.SystemColors.Control;
            resources.ApplyResources(this.lbl_Information, "lbl_Information");
            this.lbl_Information.Name = "lbl_Information";
            // 
            // Fidility
            // 
            this.AcceptButton = this.btn_OK;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btn_Cancel;
            this.Controls.Add(this.lLbl_Log);
            this.Controls.Add(this.dataGrid_Fidility);
            this.Controls.Add(this.btn_Cancel);
            this.Controls.Add(this.btn_OK);
            this.Controls.Add(this.lbl_Summary);
            this.Controls.Add(this.lbl_Information);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Fidility";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid_Fidility)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_Summary;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Button btn_OK;
        private System.Windows.Forms.DataGridView dataGrid_Fidility;
        private System.Windows.Forms.LinkLabel lLbl_Log;
        private System.Windows.Forms.Label lbl_Information;

    }
}