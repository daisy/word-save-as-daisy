using System.Collections.Generic;
using Daisy.SaveAsDAISY.Conversion.Pipeline;

namespace Daisy.SaveAsDAISY.Conversion
{
    public class EnumData : ParameterDataBase
    {
        private List<object> valuesList;
        private List<string> valuesNameList;
        private int selectedIndex;
        public bool IsEditable { get; set; }
        public string CustomValue { get; set; }

        public EnumData(Dictionary<string, object> itemsList, object defaultValue = null, bool isEditable = false) : base()
        {
            valuesList = new List<object>();
            valuesNameList = new List<string>();
            selectedIndex = -1;
            foreach (KeyValuePair<string, object> item in itemsList)
            {
                valuesNameList.Add(item.Key);
                valuesList.Add(item.Value);
            }
            IsEditable = isEditable;
            if (defaultValue != null) {
                if (valuesList.Contains(defaultValue)) {
                    selectedIndex = valuesList.IndexOf(defaultValue);
                } else if (IsEditable) {
                    CustomValue = defaultValue.ToString();
                    selectedIndex = -1;
                }
            }
                
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
            get { return selectedIndex > -1 ? valuesList[selectedIndex] : (IsEditable ? CustomValue : null); }
            set
            {
                if (value != null && valuesList.Contains(value))
                    SetSelectedIndexAndUpdateScript(valuesList.BinarySearch(value));
                else if (IsEditable)
                {
                    CustomValue = value.ToString();
                    selectedIndex = -1;
                } else throw new System.Exception("NotAbleToSelectItem");
            }
        }


        public string SelectedItemNiceName
        {
            get { return valuesNameList[selectedIndex]; }
            set
            {
                if (value != null && valuesNameList.Contains(value))
                    SetSelectedIndexAndUpdateScript(valuesNameList.BinarySearch(value));
                else if (IsEditable) {
                    CustomValue = value.ToString();
                    selectedIndex = -1;
                } else throw new System.Exception("NotAbleToSelectItem");
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
