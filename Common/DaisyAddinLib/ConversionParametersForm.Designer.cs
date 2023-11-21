namespace Daisy.SaveAsDAISY.Conversion
{
    partial class ConversionParametersForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConversionParametersForm));
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.UIDTextBox = new System.Windows.Forms.TextBox();
            this.PublisherInput = new System.Windows.Forms.TextBox();
            this.CreatorInput = new System.Windows.Forms.TextBox();
            this.TitleInput = new System.Windows.Forms.TextBox();
            this.DocumentPropertiesGroup = new System.Windows.Forms.GroupBox();
            this.AccessibilityCheckerLink = new System.Windows.Forms.LinkLabel();
            this.DocumentPropertiesTable = new System.Windows.Forms.TableLayoutPanel();
            this.TitleLabel = new System.Windows.Forms.Label();
            this.CreatorLabel = new System.Windows.Forms.Label();
            this.PublisherLabel = new System.Windows.Forms.Label();
            this.UIDLabel = new System.Windows.Forms.Label();
            this.SwitchAdvancedSettingsButton = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label_Creator = new System.Windows.Forms.Label();
            this.OKButton = new System.Windows.Forms.Button();
            this.ResetButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.BottomPanel = new System.Windows.Forms.Panel();
            this.AdvancedSettingsGroup = new System.Windows.Forms.GroupBox();
            this.AdvancedSettingsPanel = new System.Windows.Forms.Panel();
            this.DestinationControl = new Daisy.SaveAsDAISY.Forms.Controls.PathControl();
            this.DocumentPropertiesGroup.SuspendLayout();
            this.DocumentPropertiesTable.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.BottomPanel.SuspendLayout();
            this.AdvancedSettingsGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // UIDTextBox
            // 
            this.UIDTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.UIDTextBox, "UIDTextBox");
            this.UIDTextBox.Name = "UIDTextBox";
            this.toolTip1.SetToolTip(this.UIDTextBox, resources.GetString("UIDTextBox.ToolTip"));
            // 
            // PublisherInput
            // 
            this.PublisherInput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.PublisherInput, "PublisherInput");
            this.PublisherInput.Name = "PublisherInput";
            this.toolTip1.SetToolTip(this.PublisherInput, resources.GetString("PublisherInput.ToolTip"));
            // 
            // CreatorInput
            // 
            this.CreatorInput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.CreatorInput, "CreatorInput");
            this.CreatorInput.Name = "CreatorInput";
            this.toolTip1.SetToolTip(this.CreatorInput, resources.GetString("CreatorInput.ToolTip"));
            // 
            // TitleInput
            // 
            this.TitleInput.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.TitleInput, "TitleInput");
            this.TitleInput.Name = "TitleInput";
            this.toolTip1.SetToolTip(this.TitleInput, resources.GetString("TitleInput.ToolTip"));
            // 
            // DocumentPropertiesGroup
            // 
            resources.ApplyResources(this.DocumentPropertiesGroup, "DocumentPropertiesGroup");
            this.DocumentPropertiesGroup.BackColor = System.Drawing.SystemColors.Control;
            this.DocumentPropertiesGroup.Controls.Add(this.AccessibilityCheckerLink);
            this.DocumentPropertiesGroup.Controls.Add(this.DocumentPropertiesTable);
            this.DocumentPropertiesGroup.ForeColor = System.Drawing.SystemColors.ControlText;
            this.DocumentPropertiesGroup.Name = "DocumentPropertiesGroup";
            this.DocumentPropertiesGroup.TabStop = false;
            // 
            // AccessibilityCheckerLink
            // 
            resources.ApplyResources(this.AccessibilityCheckerLink, "AccessibilityCheckerLink");
            this.AccessibilityCheckerLink.LinkBehavior = System.Windows.Forms.LinkBehavior.AlwaysUnderline;
            this.AccessibilityCheckerLink.Name = "AccessibilityCheckerLink";
            this.AccessibilityCheckerLink.TabStop = true;
            this.AccessibilityCheckerLink.UseCompatibleTextRendering = true;
            this.AccessibilityCheckerLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.AccessibilityCheckerLink_LinkClicked);
            // 
            // DocumentPropertiesTable
            // 
            resources.ApplyResources(this.DocumentPropertiesTable, "DocumentPropertiesTable");
            this.DocumentPropertiesTable.BackColor = System.Drawing.SystemColors.Control;
            this.DocumentPropertiesTable.Controls.Add(this.TitleLabel, 0, 0);
            this.DocumentPropertiesTable.Controls.Add(this.TitleInput, 1, 0);
            this.DocumentPropertiesTable.Controls.Add(this.CreatorLabel, 0, 1);
            this.DocumentPropertiesTable.Controls.Add(this.CreatorInput, 1, 1);
            this.DocumentPropertiesTable.Controls.Add(this.PublisherLabel, 0, 2);
            this.DocumentPropertiesTable.Controls.Add(this.PublisherInput, 1, 2);
            this.DocumentPropertiesTable.Controls.Add(this.UIDLabel, 0, 3);
            this.DocumentPropertiesTable.Controls.Add(this.UIDTextBox, 1, 3);
            this.DocumentPropertiesTable.Name = "DocumentPropertiesTable";
            // 
            // TitleLabel
            // 
            resources.ApplyResources(this.TitleLabel, "TitleLabel");
            this.TitleLabel.ForeColor = System.Drawing.SystemColors.ControlText;
            this.TitleLabel.Name = "TitleLabel";
            // 
            // CreatorLabel
            // 
            resources.ApplyResources(this.CreatorLabel, "CreatorLabel");
            this.CreatorLabel.ForeColor = System.Drawing.SystemColors.ControlText;
            this.CreatorLabel.Name = "CreatorLabel";
            // 
            // PublisherLabel
            // 
            resources.ApplyResources(this.PublisherLabel, "PublisherLabel");
            this.PublisherLabel.ForeColor = System.Drawing.SystemColors.ControlText;
            this.PublisherLabel.Name = "PublisherLabel";
            // 
            // UIDLabel
            // 
            resources.ApplyResources(this.UIDLabel, "UIDLabel");
            this.UIDLabel.ForeColor = System.Drawing.SystemColors.ControlText;
            this.UIDLabel.Name = "UIDLabel";
            // 
            // SwitchAdvancedSettingsButton
            // 
            resources.ApplyResources(this.SwitchAdvancedSettingsButton, "SwitchAdvancedSettingsButton");
            this.SwitchAdvancedSettingsButton.ForeColor = System.Drawing.SystemColors.ControlText;
            this.SwitchAdvancedSettingsButton.Name = "SwitchAdvancedSettingsButton";
            this.SwitchAdvancedSettingsButton.UseVisualStyleBackColor = true;
            this.SwitchAdvancedSettingsButton.Click += new System.EventHandler(this.OnClickSwitchAdvandedSettingsButton);
            // 
            // tableLayoutPanel1
            // 
            resources.ApplyResources(this.tableLayoutPanel1, "tableLayoutPanel1");
            this.tableLayoutPanel1.Controls.Add(this.label_Creator, 0, 1);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            // 
            // label_Creator
            // 
            resources.ApplyResources(this.label_Creator, "label_Creator");
            this.label_Creator.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label_Creator.Name = "label_Creator";
            // 
            // OKButton
            // 
            resources.ApplyResources(this.OKButton, "OKButton");
            this.OKButton.BackColor = System.Drawing.SystemColors.Control;
            this.OKButton.ForeColor = System.Drawing.SystemColors.ControlText;
            this.OKButton.Name = "OKButton";
            this.OKButton.UseVisualStyleBackColor = false;
            this.OKButton.Click += new System.EventHandler(this.onClickTranslateButton);
            // 
            // ResetButton
            // 
            resources.ApplyResources(this.ResetButton, "ResetButton");
            this.ResetButton.BackColor = System.Drawing.SystemColors.Control;
            this.ResetButton.ForeColor = System.Drawing.SystemColors.ControlText;
            this.ResetButton.Name = "ResetButton";
            this.ResetButton.UseVisualStyleBackColor = false;
            this.ResetButton.Click += new System.EventHandler(this.btn_Reset_Click);
            // 
            // CancelButton
            // 
            resources.ApplyResources(this.CancelButton, "CancelButton");
            this.CancelButton.BackColor = System.Drawing.SystemColors.Control;
            this.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelButton.ForeColor = System.Drawing.SystemColors.ControlText;
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.UseVisualStyleBackColor = false;
            this.CancelButton.Click += new System.EventHandler(this.onClickCancelButton);
            // 
            // BottomPanel
            // 
            resources.ApplyResources(this.BottomPanel, "BottomPanel");
            this.BottomPanel.Controls.Add(this.SwitchAdvancedSettingsButton);
            this.BottomPanel.Controls.Add(this.OKButton);
            this.BottomPanel.Controls.Add(this.CancelButton);
            this.BottomPanel.Controls.Add(this.ResetButton);
            this.BottomPanel.Name = "BottomPanel";
            // 
            // AdvancedSettingsGroup
            // 
            resources.ApplyResources(this.AdvancedSettingsGroup, "AdvancedSettingsGroup");
            this.AdvancedSettingsGroup.Controls.Add(this.AdvancedSettingsPanel);
            this.AdvancedSettingsGroup.Name = "AdvancedSettingsGroup";
            this.AdvancedSettingsGroup.TabStop = false;
            // 
            // AdvancedSettingsPanel
            // 
            resources.ApplyResources(this.AdvancedSettingsPanel, "AdvancedSettingsPanel");
            this.AdvancedSettingsPanel.Name = "AdvancedSettingsPanel";
            // 
            // DestinationControl
            // 
            resources.ApplyResources(this.DestinationControl, "DestinationControl");
            this.DestinationControl.BackColor = System.Drawing.Color.Transparent;
            this.DestinationControl.Name = "DestinationControl";
            this.DestinationControl.ParameterName = "Destination:";
            this.DestinationControl.SelectedPath = null;
            // 
            // ConversionParametersForm
            // 
            this.AcceptButton = this.OKButton;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoValidate = System.Windows.Forms.AutoValidate.EnablePreventFocusChange;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.AdvancedSettingsGroup);
            this.Controls.Add(this.BottomPanel);
            this.Controls.Add(this.DocumentPropertiesGroup);
            this.Controls.Add(this.DestinationControl);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ConversionParametersForm";
            this.ShowIcon = false;
            this.Load += new System.EventHandler(this.ConversionParametersFrom_Load);
            this.DocumentPropertiesGroup.ResumeLayout(false);
            this.DocumentPropertiesGroup.PerformLayout();
            this.DocumentPropertiesTable.ResumeLayout(false);
            this.DocumentPropertiesTable.PerformLayout();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.BottomPanel.ResumeLayout(false);
            this.AdvancedSettingsGroup.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.GroupBox DocumentPropertiesGroup;
        private System.Windows.Forms.Button SwitchAdvancedSettingsButton;
        private System.Windows.Forms.TableLayoutPanel DocumentPropertiesTable;
        private System.Windows.Forms.TextBox UIDTextBox;
        private System.Windows.Forms.Label CreatorLabel;
        private System.Windows.Forms.TextBox PublisherInput;
        private System.Windows.Forms.TextBox CreatorInput;
        private System.Windows.Forms.TextBox TitleInput;
        private System.Windows.Forms.Label TitleLabel;
        private System.Windows.Forms.Label UIDLabel;
        private System.Windows.Forms.Label PublisherLabel;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        //private System.Windows.Forms.TextBox tBx_Uid;
        private System.Windows.Forms.Label label_Creator;
        private System.Windows.Forms.Button ResetButton;
        private System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.Button OKButton;
        private System.Windows.Forms.LinkLabel AccessibilityCheckerLink;
        private System.Windows.Forms.Panel BottomPanel;
        private System.Windows.Forms.GroupBox AdvancedSettingsGroup;
        private Forms.Controls.PathControl DestinationControl;
        private System.Windows.Forms.Panel AdvancedSettingsPanel;
        //private System.Windows.Forms.TextBox tBx_Publisher;



    }
}