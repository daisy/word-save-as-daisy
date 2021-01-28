using System;
using System.Collections;
using System.Windows.Forms;

namespace Daisy.DaisyConverter.DaisyConverterLib.Converters
{
	/// <summary>
	/// Implements convertation with UI messages/dialogs.
	/// </summary>
    public class SingleConverterUI : SingleConverter
    {
        public SingleConverterUI(AbstractConverter converter) 
			: base(converter)
        {
        }

        public SingleConverterUI(AbstractConverter converter, ScriptParser scriptToExecute) 
			: base(converter, scriptToExecute)
        {
        }

        #region Overrides of SingleConverter

        protected override void OnValidationError(string error, string inputFile, string outputFile)
        {
            Validation validationDialog = new Validation("FailedLabel", error, inputFile, outputFile, ResManager);
            validationDialog.ShowDialog();
        }

        protected override void OnLostElements(string inputFile, string outputFile, ArrayList elements)
        {
            Fidility fidilityDialog = new Fidility("FeedbackLabel", elements, inputFile, outputFile, ResManager);
            fidilityDialog.ShowDialog();
        }

        protected override bool IsContinueDTBookGenerationOnLostElements()
        {
            DialogResult continueDTBookGenerationResult = MessageBox.Show("Do you want to create audio file", "SaveAsDAISY", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            return continueDTBookGenerationResult == DialogResult.Yes;
        }

        protected override void OnSuccess()
        {
            MessageBox.Show(ResManager.GetString("SucessLabel"), "SaveAsDAISY - Success", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        protected override void OnMasterSubValidationError(string error)
        {
            MasterSubValidation infoBox = new MasterSubValidation(error, "Validation");
            infoBox.ShowDialog();
        }

        protected override void OnSuccessMasterSubValidation(string message)
        {
            MasterSubValidation infoBox = new MasterSubValidation(message, "Success");
            infoBox.ShowDialog();
        }

        protected override void OnUnknownError(string error)
        {
            MessageBox.Show(error, "SaveAsDAISY", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        protected override void OnUnknownError(string title, string details)
        {
            InfoBox infoBox = new InfoBox(title, details, ResManager);
            infoBox.ShowDialog();
        }


        #endregion
    }
}