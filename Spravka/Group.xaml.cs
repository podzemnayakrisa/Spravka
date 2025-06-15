using System;
using System.Collections.ObjectModel;
using System.Net.Http;
using System.Windows;
using Newtonsoft.Json;
using System.Diagnostics;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using HtmlAgilityPack;
using System.Net;
using Spravka;
using Spravka.Models;
using System.Linq;
using GroupItem = Spravka.Models.GroupItem;
using System.Threading;
using System.Text.RegularExpressions;
using System.IO;

namespace Spravka
{
    public partial class Group : Window
    {
        private ObservableCollection<GroupItem> _groups = new ObservableCollection<GroupItem>();
        private const string GoogleScriptUrl = "https://script.google.com/macros/s/AKfycbwmphanUtB6Hk8-7rc8yyYHHCNjrtkywwoqOreEkOA8rWqpH6Tug8tygusoX-l93NEWHQ/exec";

        public Group()
        {
            InitializeComponent();
            GroupsDataGrid.ItemsSource = _groups;
            Loaded += async (s, e) => await LoadDataAsync();
        }

        private async Task LoadDataAsync()
        {
            try
            {
                var groups = await LoadGroupsDataAsync();
                _groups.Clear();
                foreach (var g in groups) _groups.Add(g);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки: {ex.Message}");
            }
        }

        private async Task<List<GroupItem>> LoadGroupsDataAsync()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
                    client.Timeout = TimeSpan.FromSeconds(30); // Увеличьте при необходимости

                    var cts = new CancellationTokenSource();
                    cts.CancelAfter(TimeSpan.FromSeconds(45)); // Увеличенный таймаут

                    string url = $"{GoogleScriptUrl}?action=get_groups&t={DateTime.Now.Ticks}";
                    var response = await client.GetAsync(url, cts.Token); // Передаём токен отмены

                    response.EnsureSuccessStatusCode(); // Выбросит исключение при ошибке HTTP

                    string json = await response.Content.ReadAsStringAsync();
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string logPath = $"{desktopPath}\\groups_debug.json";
                    File.WriteAllText(logPath, json);
                    Debug.WriteLine("Response saved to: " + logPath);
                    if (string.IsNullOrWhiteSpace(json))
                        throw new Exception("Пустой ответ сервера");

                    var scriptResponse = JsonConvert.DeserializeObject<GoogleScriptResponse<List<Dictionary<string, object>>>>(json);
                    return ProcessGroupData(scriptResponse?.Data);
                }
            }
            catch (OperationCanceledException)
            {
                MessageBox.Show("Загрузка групп отменена по таймауту", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Warning);
                return new List<GroupItem>();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки групп: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<GroupItem>();
            }
        }

        private List<GroupItem> ProcessGroupData(List<Dictionary<string, object>> data)
        {
            if (data == null) return new List<GroupItem>();

            return data.Skip(1) // Пропускаем заголовки
                      .Where(item => item != null)
                      .Select(item => new GroupItem
                      {
                          Name = GetValue(item, "Название"),
                          StartDate = ParseDate(GetValue(item, "Дата начала обучения")),
                          EndDate = ParseDate(GetValue(item, "Дата окончания обучения"))
                      })
                      .Where(g => !string.IsNullOrEmpty(g.Name))
                      .ToList();
        }


        private string GetValue(Dictionary<string, object> dict, string key, string defaultValue = null)
        {
            if (dict.TryGetValue(key, out var value) && value != null)
            {
                string strValue = value.ToString();
                return string.IsNullOrWhiteSpace(strValue) ? defaultValue : strValue;
            }
            return defaultValue;
        }

        private DateTime ParseDate(string dateString)
        {
            if (string.IsNullOrWhiteSpace(dateString))
                return DateTime.Now;

            string[] formats = {
        "dd.MM.yyyy",
        "yyyy-MM-dd",
        "M/d/yyyy",
        "yyyy/MM/dd",
        "dd-MM-yyyy"
    };

            if (DateTime.TryParseExact(dateString, formats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out DateTime result))
            {
                return result;
            }

            // Попробуем парсить как дату в формате timestamp (если из Google Sheets приходит число)
            if (double.TryParse(dateString, out double timestamp))
            {
                return DateTime.FromOADate(timestamp);
            }

            return DateTime.Now;
        }

        private async void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var groups = await LoadGroupsDataAsync();
                _groups = new ObservableCollection<GroupItem>(groups);
                GroupsDataGrid.ItemsSource = _groups;
                MessageBox.Show("Данные обновлены", "Успех",
                              MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении: {ex.Message}", "Ошибка",
                              MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new GroupEditDialog();
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    await AddGroupToGoogleSheet(dialog.GroupItem);
                    _groups.Add(dialog.GroupItem);
                    MessageBox.Show("Группа добавлена", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при добавлении группы: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async Task AddGroupToGoogleSheet(GroupItem item)
        {
            using (var client = new HttpClient())
            {
                var data = new
                {
                    action = "add_group",
                    name = item.Name,
                    start_date = item.StartDate.ToString("yyyy-MM-dd"),
                    end_date = item.EndDate.ToString("yyyy-MM-dd")
                };

                var content = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                var response = await client.PostAsync(GoogleScriptUrl, content);

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception(await response.Content.ReadAsStringAsync());
                }
            }
        }

        private async void EditButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsDataGrid.SelectedItem is GroupItem selectedItem)
            {
                var dialog = new GroupEditDialog(selectedItem);
                if (dialog.ShowDialog() == true)
                {
                    try
                    {
                        await UpdateGroupInGoogleSheet(dialog.GroupItem);
                        var index = _groups.IndexOf(selectedItem);
                        _groups[index] = dialog.GroupItem;
                        MessageBox.Show("Группа обновлена", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при обновлении группы: {ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите группу для редактирования", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async Task UpdateGroupInGoogleSheet(GroupItem item)
        {
            using (var client = new HttpClient())
            {
                var data = new
                {
                    action = "update_group",
                    name = item.Name,
                    start_date = item.StartDate.ToString("yyyy-MM-dd"),
                    end_date = item.EndDate.ToString("yyyy-MM-dd")
                };

                var content = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                var response = await client.PostAsync(GoogleScriptUrl, content);

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception(await response.Content.ReadAsStringAsync());
                }
            }
        }

        private async void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (GroupsDataGrid.SelectedItem is GroupItem selectedItem)
            {
                if (MessageBox.Show($"Удалить группу:\n{selectedItem.Name}?", "Подтверждение",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        await DeleteGroupFromGoogleSheet(selectedItem.Name);
                        _groups.Remove(selectedItem);
                        MessageBox.Show("Группа удалена", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при удалении группы: {ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите группу для удаления", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async Task DeleteGroupFromGoogleSheet(string groupName)
        {
            using (var client = new HttpClient())
            {
                var data = new
                {
                    action = "delete_group",
                    name = groupName
                };

                var content = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                var response = await client.PostAsync(GoogleScriptUrl, content);

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception(await response.Content.ReadAsStringAsync());
                }
            }
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Реализация поиска по названию группы
            var searchText = SearchBox.Text.ToLower();
            GroupsDataGrid.Items.Filter = item =>
            {
                var group = item as GroupItem;
                return group.Name.ToLower().Contains(searchText);
            };
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            SearchBox_TextChanged(sender, null);
        }
        
    }
    public partial class Group : Window
    {
        private readonly MainWindow _mainWindow;

        public Group(MainWindow mainWindow)
        {
            InitializeComponent();
            _mainWindow = mainWindow;
            this.Closed += (s, e) => _mainWindow?.Show();
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            _mainWindow?.Show();
            this.Close();
        }
    }
}