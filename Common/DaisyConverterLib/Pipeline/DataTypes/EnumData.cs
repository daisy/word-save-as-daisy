using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class EnumData : ParameterDataBase
    {
        private List<object> valuesList;
        private List<string> valuesNameList;
        private int selectedIndex;

        public EnumData(Dictionary<string, object> itemsList, object defaultValue = null) : base()
        {
            valuesList = new List<object>();
            valuesNameList = new List<string>();
            selectedIndex = -1;
            foreach (KeyValuePair<string, object> item in itemsList)
            {
                valuesNameList.Add(item.Key);
                valuesList.Add(item.Value);
            }

            if (defaultValue != null && valuesList.Contains(defaultValue))
                selectedIndex = valuesList.IndexOf(defaultValue);
        }

        public List<object> Values { get { return valuesList; } }
        public List<string> Keys { get { return valuesNameList; } }

        /// <summary>
        ///  Gets and sets the index of value selected.
        /// </summary>
        public int SelectedIndex
        {
            get { return selectedIndex; }
            set
            {
                if (value >= 0 && value < valuesList.Count)
                    SetSelectedIndexAndUpdateScript(value);
                else throw new System.Exception("IndexNotInRange");
            }
        }


        public object SelectedItemValue
        {
            get { return selectedIndex > -1 ? valuesList[selectedIndex] : null; }
            set
            {
                if (value != null && valuesList.Contains(value))
                    SetSelectedIndexAndUpdateScript(valuesList.BinarySearch(value));
                else throw new System.Exception("NotAbleToSelectItem");
            }
        }


        public string SelectedItemNiceName
        {
            get { return valuesNameList[selectedIndex]; }
            set
            {
                if (value != null && valuesNameList.Contains(value))
                    SetSelectedIndexAndUpdateScript(valuesNameList.BinarySearch(value));
                else throw new System.Exception("NotAbleToSelectItem");
            }
        }

        public override object Value
        {
            get => SelectedItemValue;
            set
            {
                SelectedItemValue = value;
            }
        }

        private bool SetSelectedIndexAndUpdateScript(int index)
        {
            if (index >= 0 && index < valuesList.Count)
            {
                selectedIndex = index;
                return true;
            }
            else
                return false;
        }
    }
}
