namespace Daisy.SaveAsDAISY.Forms.Controls
{
    partial class StrUserControl
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(StrUserControl));
            this.parameterInput = new System.Windows.Forms.TextBox();
            this.InputPanel = new System.Windows.Forms.Panel();
            this.InputPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // parameterNiceName
            // 
            resources.ApplyResources(this.parameterNiceName, "parameterNiceName");
            // 
            // parameterInput
            // 
            resources.ApplyResources(this.parameterInput, "parameterInput");
            this.parameterInput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.parameterInput.Name = "parameterInput";
            // 
            // InputPanel
            // 
            resources.ApplyResources(this.InputPanel, "InputPanel");
            this.InputPanel.Controls.Add(this.parameterInput);
            this.InputPanel.Name = "InputPanel";
            // 
            // StrUserControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            resources.ApplyResources(this, "$this");
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.InputPanel);
            this.Name = "StrUserControl";
            this.Controls.SetChildIndex(this.parameterNiceName, 0);
            this.Controls.SetChildIndex(this.InputPanel, 0);
            this.InputPanel.ResumeLayout(false);
            this.InputPanel.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox parameterInput;
        private System.Windows.Forms.Panel InputPanel;
    }
}
