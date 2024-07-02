namespace Daisy.SaveAsDAISY.Addins.Word2007
{
    partial class SuggestedReferences
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SuggestedReferences));
            this.button1 = new System.Windows.Forms.Button();
            this.bttncancle = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.SuggstRefdataGridView = new System.Windows.Forms.DataGridView();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SuggstRefdataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // bttncancle
            // 
            resources.ApplyResources(this.bttncancle, "bttncancle");
            this.bttncancle.Name = "bttncancle";
            this.bttncancle.UseVisualStyleBackColor = true;
            this.bttncancle.Click += new System.EventHandler(this.bttncancle_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.richTextBox1);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // richTextBox1
            // 
            resources.ApplyResources(this.richTextBox1, "richTextBox1");
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.SuggstRefdataGridView);
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // SuggstRefdataGridView
            // 
            this.SuggstRefdataGridView.AllowUserToAddRows = false;
            this.SuggstRefdataGridView.AllowUserToDeleteRows = false;
            this.SuggstRefdataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            resources.ApplyResources(this.SuggstRefdataGridView, "SuggstRefdataGridView");
            this.SuggstRefdataGridView.MultiSelect = false;
            this.SuggstRefdataGridView.Name = "SuggstRefdataGridView";
            this.SuggstRefdataGridView.RowHeadersVisible = false;
            this.SuggstRefdataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.SuggstRefdataGridView.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.SuggstRefdataGridView_CellClick);
            this.SuggstRefdataGridView.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.SuggstRefdataGridView_CellEnter);
            this.SuggstRefdataGridView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.SuggstRefdataGridView_CellContentClick);
            // 
            // SuggestedReferences
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.bttncancle);
            this.Controls.Add(this.button1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SuggestedReferences";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.SuggstRefdataGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button bttncancle;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.DataGridView SuggstRefdataGridView;
    }
}