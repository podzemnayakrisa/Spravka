using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using HtmlAgilityPack;
using System.Net;
using Spravka;
using Newtonsoft.Json;
using System.Net.Http;
using System.Globalization;
using Spravka.Models;
using System.Diagnostics;
using GroupItem = Spravka.Models.GroupItem;
using System.IO; 
using Xceed.Document.NET;
using System.Drawing; // Для Color

// Алиасы для типов из Xceed.Document.NET
using XceedBorder = Xceed.Document.NET.Border;
using TableBorderType = Xceed.Document.NET.TableBorderType;
using BorderStyle = Xceed.Document.NET.BorderStyle;
using Alignment = Xceed.Document.NET.Alignment;
using TableCellBorderType = Xceed.Document.NET.TableCellBorderType;
using Xceed.Words.NET;

namespace Spravka
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<ResponseItem> _responses;
        private ObservableCollection<ResponseItem> _allResponses;
        private FlowDocument _currentCertificate;
        private const string GoogleScriptUrl = "https://script.google.com/macros/s/AKfycbxYzvHNfsXlB2PaUVZF34Yx6RMaaxb3L93-l7GDKt6ObwLVpAVFoqhvtv5AkQ8FF6DxLA/exec";
        public MainWindow()
        {
            InitializeComponent();
            _responses = new ObservableCollection<ResponseItem>();
            _allResponses = new ObservableCollection<ResponseItem>();
            ResponsesDataGrid.ItemsSource = _responses;
            LoadDataFromGoogleSheetsAsync().ConfigureAwait(false);
        }

        private async Task LoadDataFromGoogleSheetsAsync()
        {
            try
            {
                _responses.Clear();
                _allResponses.Clear();

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
                    client.Timeout = TimeSpan.FromSeconds(30);

                    string url = $"{GoogleScriptUrl}?action=get_responses&t={DateTime.Now.Ticks}";
                    var response = await client.GetAsync(url);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception($"Ошибка сервера: {response.StatusCode}");
                    }

                    string json = await response.Content.ReadAsStringAsync();
                    Debug.WriteLine($"Полученные данные: {json}");

                    if (string.IsNullOrWhiteSpace(json))
                    {
                        throw new Exception("Получен пустой ответ от сервера");
                    }

                    var scriptResponse = JsonConvert.DeserializeObject<GoogleScriptResponse<List<Dictionary<string, object>>>>(json)
                        ?? throw new Exception("Не удалось десериализовать данные");

                    if (!scriptResponse.Success)
                    {
                        throw new Exception("Сервер вернул ошибку: " + json);
                    }

                    foreach (var item in scriptResponse.Data ?? new List<Dictionary<string, object>>())
                    {
                        try
                        {
                            var newItem = new ResponseItem
                            {
                                FullName = $"{GetValue(item, "Фамилия")} {GetValue(item, "Имя")} {GetValue(item, "Отчество")}".Trim(),
                                Email = GetValue(item, "Почта"),
                                RequestDate = ParseDate(GetValue(item, "Отметка времени")),
                                Course = GetValue(item, "Курс") ?? "Не указано",
                                EducationForm = GetValue(item, "Форма") ?? "Не указано",
                                Basis = GetValue(item, "Основа") ?? "Не указано",
                                Status = GetValue(item, "Статус") ?? "Новый",
                                Group = GetValue(item, "Группа") ?? "Не указана"
                            };

                            _responses.Add(newItem);
                            _allResponses.Add(newItem);
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Ошибка обработки элемента: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
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

            if (double.TryParse(dateString, out double timestamp))
            {
                return DateTime.FromOADate(timestamp);
            }

            return DateTime.Now;
        }

        private async Task ApplyFiltersAndSearch()
        {
            try
            {
                if (_allResponses == null) return;

                string searchText = SearchBox?.Text?.ToLower() ?? "";
                var selectedFilter = (FilterCombo.SelectedItem as ComboBoxItem)?.Content.ToString();
                var filtered = _allResponses.ToList();

                if (!string.IsNullOrWhiteSpace(searchText))
                {
                    filtered = filtered.Where(item =>
                        (item.FullName?.ToLower() ?? "").Contains(searchText) ||
                        (item.Email?.ToLower() ?? "").Contains(searchText) ||
                        (item.Status?.ToLower() ?? "").Contains(searchText))
                    .ToList();
                }

                if (selectedFilter == "Готовые")
                    filtered = filtered.Where(item => item.IsReady).ToList();
                else if (selectedFilter == "Не готовые")
                    filtered = filtered.Where(item => !item.IsReady).ToList();

                _responses.Clear();
                foreach (var item in filtered)
                {
                    _responses.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при поиске: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            if (ResponsesDataGrid.SelectedItem is ResponseItem selectedItem)
            {
                selectedItem.Status = selectedItem.IsReady ? "Готово" : "В работе";
                await UpdateStatusInGoogleSheet(selectedItem.Email, selectedItem.Status);
            }
        }

        private async Task UpdateStatusInGoogleSheet(string email, string status)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var data = new { action = "update_status", email, status };
                    var content = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                    var response = await client.PostAsync(GoogleScriptUrl, content);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении статуса: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (ResponsesDataGrid.SelectedItem is ResponseItem selectedItem)
            {
                if (MessageBox.Show($"Удалить запись:\n{selectedItem.FullName}?", "Подтверждение",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        await DeleteRowFromGoogleSheet(selectedItem.Email);
                        _responses.Remove(selectedItem);
                        _allResponses.Remove(selectedItem);
                        MessageBox.Show("Запись удалена", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при удалении: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите запись для удаления", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async Task DeleteRowFromGoogleSheet(string email)
        {
            using (var client = new HttpClient())
            {
                var data = new { action = "delete_row", email };
                var content = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                var response = await client.PostAsync(GoogleScriptUrl, content);

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception(await response.Content.ReadAsStringAsync());
                }
            }
        }

        private async void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new AddEditDialog();
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    await AddRowToGoogleSheet(dialog.ResponseItem);
                    _responses.Add(dialog.ResponseItem);
                    _allResponses.Add(dialog.ResponseItem);
                    MessageBox.Show("Запись добавлена", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при добавлении: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }


        private async Task AddRowToGoogleSheet(ResponseItem item)
        {
            using (var client = new HttpClient())
            {
                var nameParts = item.FullName.Split(' ');
                var data = new
                {
                    action = "add_row",
                    Фамилия = nameParts.Length > 0 ? nameParts[0] : "",
                    Имя = nameParts.Length > 1 ? nameParts[1] : "",
                    Отчество = nameParts.Length > 2 ? nameParts[2] : "",
                    Почта = item.Email,
                    Статус = item.Status,
                    Курс = item.Course,
                    Форма = item.EducationForm,
                    Основа = item.Basis
                };

                var content = new StringContent(JsonConvert.SerializeObject(data), Encoding.UTF8, "application/json");
                var response = await client.PostAsync(GoogleScriptUrl, content);

                if (!response.IsSuccessStatusCode)
                {
                    throw new Exception(await response.Content.ReadAsStringAsync());
                }
            }
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            if (ResponsesDataGrid.SelectedItem is ResponseItem selectedItem)
            {
                var dialog = new AddEditDialog(selectedItem);
                if (dialog.ShowDialog() == true)
                {
                    var index = _responses.IndexOf(selectedItem);
                    _responses[index] = dialog.ResponseItem;
                    _ = UpdateRowInGoogleSheet(dialog.ResponseItem);
                }
            }
            else
            {
                MessageBox.Show("Выберите запись для редактирования", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private async Task UpdateRowInGoogleSheet(ResponseItem item)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var nameParts = item.FullName.Split(' ');
                    var data = new
                    {
                        action = "update_row",
                        email = item.Email,
                        Фамилия = nameParts.Length > 0 ? nameParts[0] : "",
                        Имя = nameParts.Length > 1 ? nameParts[1] : "",
                        Отчество = nameParts.Length > 2 ? nameParts[2] : "",
                        Почта = item.Email,
                        Статус = item.Status,
                        Курс = item.Course == "Не указано" ? "" : item.Course,
                        Форма = item.EducationForm == "Не указано" ? "" : item.EducationForm,
                        Основа = item.Basis == "Не указано" ? "" : item.Basis
                    };

                    var content = new StringContent(JsonConvert.SerializeObject(data),
                        Encoding.UTF8, "application/json");
                    var response = await client.PostAsync(GoogleScriptUrl, content);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception(await response.Content.ReadAsStringAsync());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при обновлении: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private async Task<List<GroupItem>> GetGroupsListAsync()
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
                    client.Timeout = TimeSpan.FromSeconds(30);

                    string url = $"{GoogleScriptUrl}?action=get_groups&t={DateTime.Now.Ticks}";
                    var response = await client.GetAsync(url);

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception($"Ошибка сервера: {response.StatusCode}");
                    }

                    string json = await response.Content.ReadAsStringAsync();

                    if (string.IsNullOrWhiteSpace(json))
                    {
                        throw new Exception("Получен пустой ответ от сервера");
                    }

                    var scriptResponse = JsonConvert.DeserializeObject<GoogleScriptResponse<List<Dictionary<string, object>>>>(json)
                        ?? throw new Exception("Не удалось десериализовать данные");

                    if (!scriptResponse.Success || scriptResponse.Data == null)
                    {
                        throw new Exception("Сервер вернул пустые данные");
                    }

                    var groups = new List<GroupItem>();
                    foreach (var item in scriptResponse.Data.Skip(1))
                    {
                        try
                        {
                            if (item == null) continue;

                            var newItem = new GroupItem
                            {
                                Name = item.ContainsKey("Название") ? item["Название"]?.ToString() : "",
                                StartDate = ParseDate(item.ContainsKey("Дата начала обучения") ? item["Дата начала обучения"]?.ToString() : ""),
                                EndDate = ParseDate(item.ContainsKey("Дата окончания обучения") ? item["Дата окончания обучения"]?.ToString() : "")
                            };

                            if (!string.IsNullOrEmpty(newItem.Name))
                            {
                                groups.Add(newItem);
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"Ошибка обработки элемента: {ex.Message}");
                        }
                    }

                    return groups;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки групп: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<GroupItem>();
            }
        }
        private void CreateCertificateMenuItem_Click(object sender, RoutedEventArgs e)
        {
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Word Document|*.docx"
            };

            if (saveDialog.ShowDialog() == true)
            {
                GenerateCertificate(saveDialog.FileName);
                Process.Start("explorer.exe", $"/select,\"{saveDialog.FileName}\"");
            }
        }

        private void PrintCertificateMenuItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string tempFile = Path.Combine(Path.GetTempPath(), "spravka_temp.docx");
                GenerateCertificate(tempFile);

                Process.Start(new ProcessStartInfo
                {
                    FileName = tempFile,
                    Verb = "print"
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка печати: {ex.Message}");
            }
        }

        private void GenerateCertificate(string filePath)
        {
            using (var document = DocX.Create(filePath))
            {
                // Настройка полей (в мм)
                document.MarginLeft = 30f;
                document.MarginRight = 15f;
                document.MarginTop = 20f;
                document.MarginBottom = 20f;

                // Таблица шапки (2 колонки, 19 строк)
                var table = document.AddTable(19, 2);
                table.SetWidths(new float[] { 320f, 100f });
                table.Alignment = Alignment.center;

                // Заполнение левой колонки
                string[] leftContent =
                {
                    "Министерство образования Ярославской",
                    "области",
                    "",
                    "государственное профессиональное",
                    "образовательное",
                    "",
                    "автономное учреждение Ярославской",
                    "области",
                    "",
                    "\"Ярославский промышленно-экономический",
                    "колледж",
                    "",
                    "им. Н. П. Пастухова\"",
                    "",
                    "150046, г. Ярославль, ул. Гагарина, 8",
                    "",
                    "44-26-77",
                    "",
                    "« 11 » июня 2025 г",
                    "№ 6"
                };

                for (int i = 0; i < 19; i++)
                {
                    var cell = table.Rows[i].Cells[0];
                    cell.Paragraphs.First().Append(leftContent[i])
                       .Font("Times New Roman")
                       .FontSize(10)
                       .Alignment = Alignment.left;
                }

                // Создание границы с использованием System.Drawing.Color
                var border = new XceedBorder(
                    BorderStyle.Tcbs_single,
                    BorderSize.one,
                    0,
                    System.Drawing.Color.Black);

                table.SetBorder(TableBorderType.Top, border);
                table.SetBorder(TableBorderType.Bottom, border);
                table.SetBorder(TableBorderType.Left, border);
                table.SetBorder(TableBorderType.Right, border);
                table.SetBorder(TableBorderType.InsideV, border);

                // Дополнительные границы для разделителя
                table.Rows[17].Cells[0].SetBorder(TableCellBorderType.Bottom, border);
                table.Rows[17].Cells[1].SetBorder(TableCellBorderType.Bottom, border);

                document.InsertTable(table);

                // Заголовок "СПРАВКА"
                var title = document.InsertParagraph("СПРАВКА");
                title.Alignment = Alignment.center;
                title.Font("Times New Roman").FontSize(14).Bold();
                title.SpacingAfter(20f);

                // Основной текст
                var content = new[]
                {
                    "Выдана настоящая Комарову Ивану Сергеевичу",
                    "в том, что он(а) обучается в ГПОАУ ЯО «Ярославский\nпромышленно-экономический колледж",
                    "им. Н. П. Пастухова» на 2 курсе по очной форме обучения, на бюджетной\nоснове.",
                    "Выдана для предъявления по месту требования.",
                    "Срок обучения с 01.09.2021 по 30.06.2025."
                };

                foreach (var text in content)
                {
                    var p = document.InsertParagraph(text);
                    p.Font("Times New Roman").FontSize(11);
                    p.Alignment = Alignment.both;
                    p.SpacingAfter(5f);
                }

                document.InsertParagraph().SpacingAfter(20f);

                // Подпись
                var signTable = document.AddTable(1, 2);
                signTable.SetWidths(new float[] { 250f, 250f });
                signTable.Alignment = Alignment.center;

                signTable.Rows[0].Cells[0].Paragraphs.First()
                    .Append("Заведующий отделением")
                    .Font("Times New Roman").FontSize(11);

                signTable.Rows[0].Cells[1].Paragraphs.First()
                    .Append("Ю. В. Маянцева")
                    .Font("Times New Roman").FontSize(11)
                    .Alignment = Alignment.right;

                document.InsertTable(signTable);

                // Телефон
                var phone = document.InsertParagraph("Тел. 48-05-24");
                phone.Font("Times New Roman").FontSize(11);
                phone.Alignment = Alignment.left;

                document.Save();
            }
        }
    





private void OpenGroupsWindow_Click(object sender, RoutedEventArgs e)
        {
            var groupsWindow = new Group(this);
            groupsWindow.Show();
            this.Hide();
        }

        private async void RefreshButton_Click(object sender, RoutedEventArgs e) => await LoadDataFromGoogleSheetsAsync();
        private async void SearchButton_Click(object sender, RoutedEventArgs e) => await ApplyFiltersAndSearch();
        private void FilterCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) => _ = ApplyFiltersAndSearch();
        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) => _ = ApplyFiltersAndSearch();
    }
}

 

      