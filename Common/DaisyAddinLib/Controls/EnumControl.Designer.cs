namespace Daisy.SaveAsDAISY.Forms.Controls
{
    partial class EnumControl
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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EnumControl));
            this.ValueSelector = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // parameterNiceName
            // 
            resources.ApplyResources(this.parameterNiceName, "parameterNiceName");
            // 
            // ValueSelector
            // 
            this.ValueSelector.AccessibleRole = System.Windows.Forms.AccessibleRole.ComboBox;
            this.ValueSelector.AllowDrop = true;
            resources.ApplyResources(this.ValueSelector, "ValueSelector");
            this.ValueSelector.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ValueSelector.FormattingEnabled = true;
            this.ValueSelector.Name = "ValueSelector";
            this.ValueSelector.SelectedIndexChanged += new System.EventHandler(this.ValueSelector_SelectedIndexChanged);
            // 
            // EnumControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            resources.ApplyResources(this, "$this");
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.ValueSelector);
            this.Name = "EnumControl";
            this.Controls.SetChildIndex(this.ValueSelector, 0);
            this.Controls.SetChildIndex(this.parameterNiceName, 0);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox ValueSelector;
    }



}
