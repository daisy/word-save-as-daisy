using Daisy.SaveAsDAISY.Conversion;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


namespace Daisy.SaveAsDAISY.Forms.Controls
{
    /// <summary>
    /// Base of custom user controlers
    /// </summary>
    public partial class BaseUserControl : UserControl
    {

        protected BaseUserControl()
        {
            InitializeComponent();
        }

        protected ScriptParameter _linkedParameter = null;

        public ScriptParameter getLinkedParameter() => this._linkedParameter;
        public virtual void setLinkedParameter(ScriptParameter s)
        {
            this._linkedParameter = s;
            parameterNiceName.Text = s.NiceName;
            parameterNiceName.AccessibleName = s.NiceName;
        }


        public string ParameterName
        {
            get
            {
                return this.parameterNiceName.Text;
            }

            set
            {
                this.parameterNiceName.Text = value;
            }
        }

        public virtual void setWidth(int width)
        {

        }


        protected BaseUserControl(ScriptParameter p) : this()
        {
        }

        public virtual void UpdateScriptParameterValue() { }

        public int LabelWidth { get { return parameterNiceName.Width; } }

    }

}
