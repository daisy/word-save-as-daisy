using System.ComponentModel;

namespace Daisy.SaveAsDAISY.Forms.Controls
{
    partial class BaseUserControl
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
            this.components = new System.ComponentModel.Container();
            this.descriptionTooltip = new System.Windows.Forms.ToolTip(this.components);
            this.parameterNiceName = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // parameterNiceName
            // 
            this.parameterNiceName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.parameterNiceName.AutoSize = true;
            this.parameterNiceName.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.parameterNiceName.Location = new System.Drawing.Point(-3, 0);
            this.parameterNiceName.Name = "parameterNiceName";
            this.parameterNiceName.Size = new System.Drawing.Size(42, 20);
            this.parameterNiceName.TabIndex = 3;
            this.parameterNiceName.Text = "Nice:";
            // 
            // BaseUserControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Inherit;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.Color.Transparent;
            this.Controls.Add(this.parameterNiceName);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "BaseUserControl";
            this.Size = new System.Drawing.Size(42, 44);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        protected System.Windows.Forms.ToolTip descriptionTooltip;
        public System.Windows.Forms.Label parameterNiceName;

        #endregion
        //protected System.Windows.Forms.Label mLabel;


    }

}
