using Daisy.SaveAsDAISY.Conversion.Pipeline;
using Daisy.SaveAsDAISY.Conversion.Pipeline.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Daisy.SaveAsDAISY.WPF
{
    public class ModifiableProperty
    {
        public string Name { get; set; }
        public string Desc { get; set; }

        public string InitialValue { get; set; }

        private string _value;
        public string Value { 
            get => _value;
            set {
                _value = value;
            }
        }

        public ModifiableProperty(string name, string desc, string defaultValue, string value)
        {
            Name = name;
            Desc = desc;
            InitialValue = defaultValue;
            Value = value == null ? InitialValue : value;
        }
    }

    /// <summary>
    /// 
    /// </summary>
    public partial class Pipeline2Properties : Window
    {

        public List<ModifiableProperty> Properties { get; set; }

        public List<EngineProperty> OriginalProperties { get; set; }

        public Pipeline2Properties(List<EngineProperty> properties, List<EngineProperty> userSetProperties = null)
        {
            OriginalProperties = properties;
            Properties = properties.Select(
                p => new ModifiableProperty(p.Name, p.Desc, p.Value, userSetProperties?.FirstOrDefault(i => i.Name == p.Name)?.Value)
            ).ToList();
            InitializeComponent();
            PropertiesListBox.Items.Clear();
            PropertiesListBox.ItemsSource = Properties;

        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        public List<EngineProperty> UpdatedProperties { get {
                List<EngineProperty> newUserProperties = new List<EngineProperty>();
                foreach (ModifiableProperty p in Properties)
                {
                    // A value has been set and it's different from the default value, so we want to save it as a user property
                    if (p.Value != p.InitialValue)
                    {
                        newUserProperties.Add(new EngineProperty
                        {
                            Name = p.Name,
                            Desc = p.Desc,
                            Value = p.Value,
                            Href = OriginalProperties.FirstOrDefault(op => op.Name == p.Name)?.Href
                        });
                    }
                }
                return newUserProperties;
            } }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            //List<EngineProperty> newUserProperties = new List<EngineProperty>();
            //foreach (ModifiableProperty p in Properties)
            //{
            //    // A value has been set and it's different from the default value, so we want to save it as a user property
            //    if (p.Value != p.InitialValue)
            //    {
            //        newUserProperties.Add(new EngineProperty
            //        {
            //            Name = p.Name,
            //            Desc = p.Desc,
            //            Value = p.Value,
            //            Href = OriginalProperties.FirstOrDefault(op => op.Name == p.Name)?.Href
            //        });
            //    }
            //}
            //PipelineUserProperties.Instance.UpdateOrAddRange(newUserProperties);
            //PipelineUserProperties.Instance.Save();
            DialogResult = true;
            Close();
        }
    }
}
