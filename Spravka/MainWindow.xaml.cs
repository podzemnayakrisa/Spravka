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
        private CertificateGenerator _certificateGenerator = new CertificateGenerator();
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
        //private async void CreateCertificateMenuItem_Click(object sender, RoutedEventArgs e)
        // {
        //  if (ResponsesDataGrid.SelectedItem is ResponseItem student)
        //  {
        //     try
        //    {
        // using (var client = new HttpClient())
        //  {
        // client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");
        // client.Timeout = TimeSpan.FromSeconds(30);

        // Получаем данные о группе
        // var groupName = student.Group;
        // var groups = await LoadGroupsDataAsync(client);
        //  var group = groups.FirstOrDefault(g => g.Name == groupName);
        //            if (group == null)
        //           {
        // MessageBox.Show("Группа не найдена", "Ошибка",
        //   MessageBoxButton.OK, MessageBoxImage.Error);
        //                      return;
        //
        //   var requestData = new
        //     {
        //  action = "create_certificate",
        //  fullName = student.FullName,
        // course = student.Course,
        //educationForm = student.EducationForm,
        // basis = student.Basis,
        // group = groupName,
        //  startDate = group.StartDate.ToString("yyyy-MM-dd"),
        //  endDate = group.EndDate.ToString("yyyy-MM-dd")
        //   };
        // Альтернатива PostAsJsonAsync
        // var json = JsonConvert.SerializeObject(requestData);
        // var content = new StringContent(json, Encoding.UTF8, "application/json");
        //var response = await client.PostAsync(GoogleScriptUrl, content);
        //  if (!response.IsSuccessStatusCode)
        //     {
        //          throw new Exception($"Ошибка сервера: {response.StatusCode}");
        //   }
        //  var responseString = await response.Content.ReadAsStringAsync();
        //   var result = JsonConvert.DeserializeObject<Dictionary<string, string>>(responseString);

        //  Process.Start(new ProcessStartInfo(result["documentUrl"]) { UseShellExecute = true });
        //  }
        // }
        //  catch (Exception ex)
        //   {
        // MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
        // MessageBoxButton.OK, MessageBoxImage.Error);
        //        }
        // }
        //  }

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


        //private async void PrintCertificateMenuItem_Click(object sender, RoutedEventArgs e)
        // {
        //  if (ResponsesDataGrid.SelectedItem is ResponseItem selectedItem)
        // {
        //    try
        //    {
        // Получаем данные о группе  
        // var groupName = selectedItem.Group;
        //var groups = await GetGroupsListAsync(); // Используем новый метод
        // var group = groups.FirstOrDefault(g => g.Name == groupName);
        //        if (group == null)
        //    {
        // MessageBox.Show("Группа не найдена", "Ошибка",
        //  MessageBoxButton.OK, MessageBoxImage.Error);
        //          return;
        // }
        // Создаем документ справки
        // var document = CreateCertificateDocument(selectedItem, group);
        // Печать документа
        //  PrintDocument(document);

        //  MessageBox.Show("Справка отправлена на печать", "Успех",
        //  MessageBoxButton.OK, MessageBoxImage.Information);
        //  }
        //              catch (Exception ex)
        //    {
        //  MessageBox.Show($"Ошибка при печати: {ex.Message}", "Ошибка",
        //  MessageBoxButton.OK, MessageBoxImage.Error);
        //  }
        // }
        //    else
        // {
        // MessageBox.Show("Выберите запись для печати", "Внимание",
        //  MessageBoxButton.OK, MessageBoxImage.Warning);
        //  }
        //  }
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
        public class CertificateGenerator
        {
            private int _certificateCounter = 1;

            public void GenerateAndPrintCertificate(ResponseItem student, GroupItem group = null)
            {
                try
                {
                    FlowDocument certificate = CreateCertificateDocument(student, group);
                    PrintDocument(certificate);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                                  MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            private FlowDocument CreateCertificateDocument(ResponseItem student, GroupItem group = null)
            {
                var doc = new FlowDocument
                {
                    PageWidth = 794,
                    PageHeight = 1123,
                    PagePadding = new Thickness(40),
                    FontFamily = new FontFamily("Times New Roman"),
                    FontSize = 12,
                    TextAlignment = TextAlignment.Left
                };

                // 1. Шапка с логотипом и реквизитами
                AddHeaderWithLogo(doc);

                // 2. Номер справки и дата
                AddCertificateNumberAndDate(doc);

                // 3. Заголовок "СПРАВКА"
                AddCertificateTitle(doc);

                // 4. Основное содержание
                AddMainContent(doc, student, group);

                // 5. Подпись и печать
                AddSignatureBlock(doc);

                return doc;
            }

            private void AddHeaderWithLogo(FlowDocument doc)
            {
                var grid = new Table
                {
                    CellSpacing = 0,
                    Margin = new Thickness(0, 0, 0, 20)
                };

                var column1 = new TableColumn { Width = new GridLength(400) };
                var column2 = new TableColumn { Width = new GridLength(200) };
                grid.Columns.Add(column1);
                grid.Columns.Add(column2);

                var rowGroup = new TableRowGroup();
                grid.RowGroups.Add(rowGroup);

                // Добавляем строки с реквизитами
                string[] headerLines = {
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
            "(ГПОАУ ЯО \"Ярославский",
            "промышленно-экономический колледж",
            "",
            "им. Н. П. Пастухова\")",
            "",
            "150046, г. Ярославль, ул. Гагарина, 8",
            "",
            "44-26-77"
        };

                foreach (var line in headerLines)
                {
                    var row = new TableRow();
                    var cell1 = new TableCell(new Paragraph(new Run(line)));
                    var cell2 = new TableCell();

                    if (line.StartsWith("\"") || line.StartsWith("(ГПОАУ"))
                    {
                        cell1.ColumnSpan = 2;
                        row.Cells.Add(cell1);
                    }
                    else
                    {
                        row.Cells.Add(cell1);
                        row.Cells.Add(cell2);
                    }

                    rowGroup.Rows.Add(row);
                }

                doc.Blocks.Add(grid);
            }

            private void AddCertificateNumberAndDate(FlowDocument doc)
            {
                var currentDate = DateTime.Now;

                var grid = new Table
                {
                    CellSpacing = 0,
                    Margin = new Thickness(0, 0, 0, 20)
                };

                grid.Columns.Add(new TableColumn { Width = new GridLength(400) });
                grid.Columns.Add(new TableColumn { Width = new GridLength(200) });

                var rowGroup = new TableRowGroup();
                grid.RowGroups.Add(rowGroup);

                // Номер справки
                var numberRow = new TableRow();
                numberRow.Cells.Add(new TableCell(new Paragraph(new Run(""))));
                numberRow.Cells.Add(new TableCell(new Paragraph(
                    new Run($"№ {currentDate:yyyyMM}-{_certificateCounter++:000}"))
                { TextAlignment = TextAlignment.Right }));
                rowGroup.Rows.Add(numberRow);

                // Разделительная линия
                var lineRow = new TableRow();
                lineRow.Cells.Add(new TableCell(new Paragraph(
                    new Run("________________________"))));
                lineRow.Cells.Add(new TableCell());
                rowGroup.Rows.Add(lineRow);

                // Дата
                var dateRow = new TableRow();
                dateRow.Cells.Add(new TableCell(new Paragraph(
                    new Run($"_________________{currentDate:yyyy} г."))));
                dateRow.Cells.Add(new TableCell());
                rowGroup.Rows.Add(dateRow);

                doc.Blocks.Add(grid);
            }

            private void AddCertificateTitle(FlowDocument doc)
            {
                var title = new Paragraph(new Run("СПРАВКА"))
                {
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    TextAlignment = TextAlignment.Center,
                    Margin = new Thickness(0, 20, 0, 30)
                };
                doc.Blocks.Add(title);
            }

            private void AddMainContent(FlowDocument doc, ResponseItem student, GroupItem group)
            {
                var content = new Paragraph
                {
                    Margin = new Thickness(0, 0, 0, 20),
                    TextAlignment = TextAlignment.Justify
                };

                content.Inlines.Add(new Run("Выдана настоящая "));
                content.Inlines.Add(new Bold(new Run(student.FullName)));

                content.Inlines.Add(new Run(", в том, что он(а) обучается в ГПОАУ ЯО \"Ярославский\n" +
                                         "промышленно-экономический колледж\n\n" +
                                         "им. Н. П. Пастухова\" на "));
                content.Inlines.Add(new Bold(new Run(student.Course)));
                content.Inlines.Add(new Run(" курсе по "));
                content.Inlines.Add(new Bold(new Run(student.EducationForm)));
                content.Inlines.Add(new Run(" форме обучения, на "));
                content.Inlines.Add(new Bold(new Run(student.Basis)));
                content.Inlines.Add(new Run(" основе.\n\n" +
                                         "Выдана для предъявления по месту требования.\n\n"));

                if (group != null)
                {
                    content.Inlines.Add(new Run("Срок обучения с "));
                    content.Inlines.Add(new Bold(new Run(group.StartDate.ToString("dd.MM.yyyy"))));
                    content.Inlines.Add(new Run(" по "));
                    content.Inlines.Add(new Bold(new Run(group.EndDate.ToString("dd.MM.yyyy"))));
                    content.Inlines.Add(new Run("."));
                }

                doc.Blocks.Add(content);
            }

            private void AddSignatureBlock(FlowDocument doc)
            {
                var table = new Table
                {
                    CellSpacing = 0,
                    Margin = new Thickness(0, 40, 0, 0)
                };

                table.Columns.Add(new TableColumn { Width = new GridLength(300) });
                table.Columns.Add(new TableColumn { Width = new GridLength(300) });

                var rowGroup = new TableRowGroup();
                table.RowGroups.Add(rowGroup);

                // Подпись
                var signRow = new TableRow();
                signRow.Cells.Add(new TableCell(new Paragraph(new Run("Заведующий отделением"))));
                signRow.Cells.Add(new TableCell(new Paragraph(
                    new Run("_________________________"))
                { TextAlignment = TextAlignment.Right }));
                rowGroup.Rows.Add(signRow);

                // Телефон
                var phoneRow = new TableRow();
                phoneRow.Cells.Add(new TableCell(new Paragraph(new Run("Тел. 48-05-24"))));
                phoneRow.Cells.Add(new TableCell());
                rowGroup.Rows.Add(phoneRow);

                doc.Blocks.Add(table);
            }

            private void PrintDocument(FlowDocument document)
            {
                var printDialog = new PrintDialog();

                document.PageHeight = printDialog.PrintableAreaHeight;
                document.PageWidth = printDialog.PrintableAreaWidth;

                if (printDialog.ShowDialog() == true)
                {
                    IDocumentPaginatorSource paginatorSource = document;
                    printDialog.PrintDocument(paginatorSource.DocumentPaginator, "Справка студента");

                    MessageBox.Show("Справка успешно отправлена на печать",
                                  "Успешно",
                                  MessageBoxButton.OK,
                                  MessageBoxImage.Information);
                }
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