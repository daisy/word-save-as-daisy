/* 
 * Copyright (c) 2006, Clever Age
 * All rights reserved.
 * 
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of Clever Age nor the names of its contributors 
 *       may be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

namespace Sonata.DaisyConverter.DaisyConverterLib
{
    partial class ConverterForm
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConverterForm));
			this.progressBar1 = new System.Windows.Forms.ProgressBar();
			this.cancelButton = new System.Windows.Forms.Button();
			this.lblProgress = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// progressBar1
			// 
			resources.ApplyResources(this.progressBar1, "progressBar1");
            this.progressBar1.Maximum = 800;
			this.progressBar1.Name = "progressBar1";
			this.progressBar1.UseWaitCursor = true;
			// 
			// cancelButton
			// 
			this.cancelButton.BackColor = System.Drawing.SystemColors.ControlLight;
			resources.ApplyResources(this.cancelButton, "cancelButton");
			this.cancelButton.Name = "cancelButton";
			this.cancelButton.UseVisualStyleBackColor = false;
			this.cancelButton.UseWaitCursor = true;
			this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
			// 
			// lblProgress
			// 
			resources.ApplyResources(this.lblProgress, "lblProgress");
			this.lblProgress.Name = "lblProgress";
			this.lblProgress.UseWaitCursor = true;
			// 
			// ConverterForm
			// 
			resources.ApplyResources(this, "$this");
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ControlBox = false;
			this.Controls.Add(this.lblProgress);
			this.Controls.Add(this.cancelButton);
			this.Controls.Add(this.progressBar1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "ConverterForm";
			this.ShowInTaskbar = false;
			this.UseWaitCursor = true;
			this.Load += new System.EventHandler(this.ConverterForm_Load);
			this.Shown += new System.EventHandler(this.ConverterForm_Shown);
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label lblProgress;

    }
}
