using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


namespace Daisy.SaveAsDAISY.Forms.Controls
{
    public partial class BaseUserControl : UserControl
    {
        public BaseUserControl() { InitializeComponent(); }

        //public virtual String DescriptionLabel
        //{
        //   // get { return mLabel.Text; }
        //   // set { mLabel.Text = value; }
        //}

        public virtual void UpdateScriptParameterValue() { }
    }

}
