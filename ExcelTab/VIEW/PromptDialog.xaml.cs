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

namespace ExcelTab.VIEW
{
    /// <summary>
    /// Interaction logic for PromptDialog.xaml
    /// </summary>
    public partial class PromptDialog : Window
    {
        public enum InputType
        {
            Date,
            Text,
            Password,
            Number
        }

        private InputType _inputType = InputType.Text;

        private string _defaultValue;

        public PromptDialog(string question, string title, string defaultValue = "", InputType inputType = InputType.Text)
        {
            InitializeComponent();
            Loaded += new RoutedEventHandler(PromptDialog_Loaded);
            txtQuestion.Text = question;
            Title = title;
            _defaultValue = defaultValue;
            _inputType = inputType;
        }

        private void PromptDialog_Loaded(object sender, RoutedEventArgs e)
        {
            switch (_inputType)
            {
                case InputType.Password:
                    txtPasswordResponse.Visibility = Visibility.Visible;
                    txtPasswordResponse.Focus();
                    break;
                case InputType.Date:
                    dtpDateResponse.Visibility = Visibility.Visible;
                    if (!string.IsNullOrEmpty(_defaultValue))
                        dtpDateResponse.SelectedDate = DateTime.Parse(_defaultValue);

                    dtpDateResponse.Focus();
                    break;
                default:
                    txtResponse.Visibility = Visibility.Visible;
                    txtResponse.Text = _defaultValue;
                    txtResponse.Focus();
                    txtResponse.SelectAll();
                    break;
            }
        }

        public static string Prompt(string question, string title, string defaultValue = "", InputType inputType = InputType.Text)
        {
            PromptDialog inst = new PromptDialog(question, title, defaultValue, inputType);
            inst.ShowDialog();
            if (inst.DialogResult == true)
                return inst.ResponseText;
            return null;
        }

        public string ResponseText
        {
            get
            {
                switch (_inputType)
                {
                    case InputType.Password: return txtPasswordResponse.Password;
                    case InputType.Date: return dtpDateResponse.SelectedDate.ToString();
                    default: return txtResponse.Text;
                }
            }
        }

        private void btnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
