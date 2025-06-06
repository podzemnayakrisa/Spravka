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

namespace Spravka
{
    /// <summary>
    /// Логика взаимодействия для AddEditDialog.xaml
    /// </summary>
    public partial class AddEditDialog : Window
    {
        public ResponseItem ResponseItem { get; private set; }
        public AddEditDialog()
        {
            InitializeComponent();
            ResponseItem = new ResponseItem();
        }
        public AddEditDialog(ResponseItem item) : this()
        {
            ResponseItem = item;
            txtFullName.Text = item.FullName;
            txtEmail.Text = item.Email;
            txtCourse.Text = item.Course;

            // Устанавливаем значения для новых полей
            cmbEducationForm.SelectedValue = item.EducationForm;
            cmbBasis.SelectedValue = item.Basis;
            cmbStatus.SelectedValue = item.Status;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ResponseItem.FullName = txtFullName.Text;
            ResponseItem.Email = txtEmail.Text;
            ResponseItem.Course = txtCourse.Text;

            // Сохраняем новые поля
            ResponseItem.EducationForm = (cmbEducationForm.SelectedItem as ComboBoxItem)?.Content.ToString();
            ResponseItem.Basis = (cmbBasis.SelectedItem as ComboBoxItem)?.Content.ToString();
            ResponseItem.Status = (cmbStatus.SelectedItem as ComboBoxItem)?.Content.ToString();

            ResponseItem.IsReady = ResponseItem.Status == "Готово";

            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}

