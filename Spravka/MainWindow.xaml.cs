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

namespace Spravka
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private ObservableCollection<ResponseItem> _responses;
        private ObservableCollection<ResponseItem> _allResponses; // Хранит все данные для фильтрации
        private const string GoogleScriptUrl = "https://script.google.com/macros/s/AKfycbxK4jsiq0-RNG1rncdeInRABFZ8dlW2kiqhGVzlFyaLLAaZlXC3iqorEb6yRHJ9BMMGCQ/exec";
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

                    // Изменяем десериализацию
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
        private async void CreateCertificateMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (ResponsesDataGrid.SelectedItem is ResponseItem student)
            {
                try
                {
                    using (var client = new HttpClient())
                    {
                        client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
                        client.Timeout = TimeSpan.FromSeconds(30);

                        // Получаем данные о группе
                        var groupName = student.Group;
                        var groups = await LoadGroupsDataAsync(client);
                        var group = groups.FirstOrDefault(g => g.Name == groupName);

                        if (group == null)
                        {
                            MessageBox.Show("Группа не найдена", "Ошибка",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }

                        var requestData = new
                        {
                            action = "create_certificate",
                            fullName = student.FullName,
                            course = student.Course,
                            educationForm = student.EducationForm,
                            basis = student.Basis,
                            group = groupName,
                            startDate = group.StartDate.ToString("yyyy-MM-dd"),
                            endDate = group.EndDate.ToString("yyyy-MM-dd")
                        };

                        // Альтернатива PostAsJsonAsync
                        var json = JsonConvert.SerializeObject(requestData);
                        var content = new StringContent(json, Encoding.UTF8, "application/json");
                        var response = await client.PostAsync(GoogleScriptUrl, content);

                        if (!response.IsSuccessStatusCode)
                        {
                            throw new Exception($"Ошибка сервера: {response.StatusCode}");
                        }

                        var responseString = await response.Content.ReadAsStringAsync();
                        var result = JsonConvert.DeserializeObject<Dictionary<string, string>>(responseString);

                        Process.Start(new ProcessStartInfo(result["documentUrl"]) { UseShellExecute = true });
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async Task<List<GroupItem>> LoadGroupsDataAsync(HttpClient client)
        {
            try
            {
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
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки групп: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return new List<GroupItem>();
            }
        }


        private async void PrintCertificateMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (ResponsesDataGrid.SelectedItem is ResponseItem selectedItem)
            {
                try
                {
                    // Получаем данные о группе  
                    var groupName = selectedItem.Group;
                    var groups = await GetGroupsListAsync(); // Используем новый метод
                    var group = groups.FirstOrDefault(g => g.Name == groupName);

                    if (group == null)
                    {
                        MessageBox.Show("Группа не найдена", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                    // Создаем документ справки
                    var document = CreateCertificateDocument(selectedItem, group);

                    // Печать документа
                    PrintDocument(document);

                    MessageBox.Show("Справка отправлена на печать", "Успех",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при печати: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Выберите запись для печати", "Внимание",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
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

        private void PrintDocument(FlowDocument document)
        {
            var printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                document.PageHeight = printDialog.PrintableAreaHeight;
                document.PageWidth = printDialog.PrintableAreaWidth;
                document.PagePadding = new Thickness(50);

                IDocumentPaginatorSource paginatorSource = document;
                printDialog.PrintDocument(paginatorSource.DocumentPaginator, "Печать справки");
            }
        }

        private FlowDocument CreateCertificateDocument(ResponseItem student, GroupItem group = null)
        {
            // Создаем документ с оформлением
            var doc = new FlowDocument
            {
                PageWidth = 794,  // A4 в pixels (96 dpi * 210mm / 25.4)
                PageHeight = 1123,
                PagePadding = new Thickness(50),
                FontFamily = new FontFamily("Times New Roman"),
                FontSize = 12
            };

            // Шапка колледжа
            var collegeHeader = new Paragraph();
            collegeHeader.Inlines.Add(new Run("Министерство образования Ярославской области") { FontSize = 12 });
            collegeHeader.Inlines.Add(new LineBreak());
            collegeHeader.Inlines.Add(new Run("государственное профессиональное образовательное") { FontSize = 12 });
            collegeHeader.Inlines.Add(new LineBreak());
            collegeHeader.Inlines.Add(new Run("автономное учреждение Ярославской области") { FontSize = 12 });
            collegeHeader.Inlines.Add(new LineBreak());
            collegeHeader.Inlines.Add(new Run("«Ярославский промышленно-экономический колледж") { FontSize = 12 });
            collegeHeader.Inlines.Add(new LineBreak());
            collegeHeader.Inlines.Add(new Run("им. Н. П. Пастухова»") { FontSize = 12 });
            collegeHeader.Inlines.Add(new LineBreak());
            collegeHeader.Inlines.Add(new Run("(ГПОАУ ЯО «Ярославский промышленно-экономический колледж") { FontSize = 12 });
            collegeHeader.Inlines.Add(new LineBreak());
            collegeHeader.Inlines.Add(new Run("им. Н. П. Пастухова»)") { FontSize = 12 });
            collegeHeader.Inlines.Add(new LineBreak());
            collegeHeader.Inlines.Add(new Run("150046, г. Ярославль, ул. Гагарина, 8") { FontSize = 12 });
            collegeHeader.Inlines.Add(new LineBreak());
            collegeHeader.Inlines.Add(new Run("44-26-77") { FontSize = 12 });
            collegeHeader.TextAlignment = TextAlignment.Left;
            collegeHeader.Margin = new Thickness(0, 0, 0, 20);
            doc.Blocks.Add(collegeHeader);

            // Разделительная линия
            var line = new Paragraph();
            line.Inlines.Add(new Run(new string('_', 100)) { FontSize = 12 });
            line.TextAlignment = TextAlignment.Center;
            line.Margin = new Thickness(0, 0, 0, 20);
            doc.Blocks.Add(line);

            // Заголовок "СПРАВКА"
            var header = new Paragraph(new Run("СПРАВКА"))
            {
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                TextAlignment = TextAlignment.Center,
                Margin = new Thickness(0, 0, 0, 20)
            };
            doc.Blocks.Add(header);

            // Основное содержание
            var content = new Paragraph();
            content.Inlines.Add(new Run("Выдана настоящая ") { FontSize = 12 });
            content.Inlines.Add(new Run(student.FullName) { FontWeight = FontWeights.Bold, FontSize = 12 });
            content.Inlines.Add(new LineBreak());
            content.Inlines.Add(new Run("в том, что он(а) обучается в ГПОАУ ЯО «Ярославский промышленно-экономический колледж") { FontSize = 12 });
            content.Inlines.Add(new LineBreak());
            content.Inlines.Add(new Run("им. Н. П. Пастухова» на ") { FontSize = 12 });
            content.Inlines.Add(new Run(student.Course) { FontWeight = FontWeights.Bold, FontSize = 12 });
            content.Inlines.Add(new Run(" курсе по ") { FontSize = 12 });
            content.Inlines.Add(new Run(student.EducationForm) { FontWeight = FontWeights.Bold, FontSize = 12 });
            content.Inlines.Add(new Run(" форме обучения, на ") { FontSize = 12 });
            content.Inlines.Add(new Run(student.Basis) { FontWeight = FontWeights.Bold, FontSize = 12 });
            content.Inlines.Add(new Run(" основе.") { FontSize = 12 });
            content.Inlines.Add(new LineBreak());
            content.Inlines.Add(new LineBreak());
            content.Inlines.Add(new Run("Выдана для предъявления по месту требования.") { FontSize = 12 });
            content.Inlines.Add(new LineBreak());
            content.Inlines.Add(new LineBreak());

            // Добавляем срок обучения, если указана группа
            if (group != null)
            {
                content.Inlines.Add(new Run("Срок обучения с ") { FontSize = 12 });
                content.Inlines.Add(new Run(group.StartDate.ToString("dd.MM.yyyy")) { FontWeight = FontWeights.Bold, FontSize = 12 });
                content.Inlines.Add(new Run(" до ") { FontSize = 12 });
                content.Inlines.Add(new Run(group.EndDate.ToString("dd.MM.yyyy")) { FontWeight = FontWeights.Bold, FontSize = 12 });
                content.Inlines.Add(new Run(".") { FontSize = 12 });
            }

            doc.Blocks.Add(content);

            // Подпись
            var sign = new Paragraph
            {
                Margin = new Thickness(0, 40, 0, 0),
                TextAlignment = TextAlignment.Left
            };
            sign.Inlines.Add(new Run("Заведующий отделением") { FontSize = 12 });
            sign.Inlines.Add(new LineBreak());
            sign.Inlines.Add(new Run("Тел. 48-05-24") { FontSize = 12 });
            doc.Blocks.Add(sign);

            return doc;
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