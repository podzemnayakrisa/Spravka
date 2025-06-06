using System;
using System.Windows;
using GroupItem = Spravka.Models.GroupItem;

namespace Spravka
{
    public partial class GroupEditDialog : Window
    {
        public GroupItem GroupItem { get; private set; }

        public GroupEditDialog()
        {
            InitializeComponent();
            GroupItem = new GroupItem
            {
                StartDate = DateTime.Now,
                EndDate = DateTime.Now.AddYears(1)
            };
            dpStartDate.SelectedDate = GroupItem.StartDate;
            dpEndDate.SelectedDate = GroupItem.EndDate;
        }

        public GroupEditDialog(GroupItem item) : this()
        {
            GroupItem = item;
            txtName.Text = item.Name;
            dpStartDate.SelectedDate = item.StartDate;
            dpEndDate.SelectedDate = item.EndDate;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("Введите название группы", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (dpStartDate.SelectedDate == null || dpEndDate.SelectedDate == null)
            {
                MessageBox.Show("Укажите даты начала и окончания обучения", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (dpStartDate.SelectedDate > dpEndDate.SelectedDate)
            {
                MessageBox.Show("Дата начала не может быть позже даты окончания", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            GroupItem.Name = txtName.Text;
            GroupItem.StartDate = dpStartDate.SelectedDate.Value;
            GroupItem.EndDate = dpEndDate.SelectedDate.Value;

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