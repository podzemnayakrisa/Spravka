using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;
using Newtonsoft.Json;
using Spravka.Models;

namespace Spravka
{
    public partial class MainWindow : Window
    {
        private ObservableCollection<ResponseItem> _responses;
        private ObservableCollection<ResponseItem> _allResponses;
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

        private void CreateCertificateMenuItem_Click(object sender, RoutedEventArgs e)
        {
            // Простая реализация сохранения как текстового файла
            var saveDialog = new SaveFileDialog
            {
                Filter = "Текстовый файл|*.txt",
                FileName = "справка.txt"
            };

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    File.WriteAllText(saveDialog.FileName, GenerateCertificateText());
                    Process.Start("explorer.exe", $"/select,\"{saveDialog.FileName}\"");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при сохранении: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void PrintCertificateMenuItem_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                try
                {
                    // Создаем документ для печати
                    FlowDocument document = CreatePrintableFlowDocument();

                    // Печатаем документ
                    printDialog.PrintDocument(
                        ((IDocumentPaginatorSource)document).DocumentPaginator,
                        "Справка об обучении");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при печати: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private string GenerateCertificateText()
        {
            return @"
Министерство образования Ярославской области
государственное профессиональное образовательное
автономное учреждение Ярославской области
""Ярославский промышленно-экономический колледж им. Н. П. Пастухова""
150046, г. Ярославль, ул. Гагарина, 8
44-26-77

« 11 » июня 2025 г   № 6

СПРАВКА

Выдана настоящая Комарову Ивану Сергеевичу в том, что он(а) обучается 
в ГПОАУ ЯО «Ярославский промышленно-экономический колледж им. Н. П. Пастухова» 
на 2 курсе по очной форме обучения, на бюджетной основе.

Выдана для предъявления по месту требования.

Срок обучения с 01.09.2021 по 30.06.2025.

Заведующий отделением                     Ю. В. Маянцева

Тел. 48-05-24";
        }

        private FlowDocument CreatePrintableFlowDocument()
        {
            // Создаем FlowDocument с правильными отступами
            FlowDocument document = new FlowDocument();
            document.PagePadding = new Thickness(40);
            document.PageWidth = 794; // Ширина A4 в пикселях (96 dpi)

            // Шапка документа
            Paragraph header = new Paragraph();
            header.Inlines.Add(new Run(
                "Министерство образования Ярославской области\n" +
                "государственное профессиональное образовательное\n" +
                "автономное учреждение Ярославской области\n" +
                "\"Ярославский промышленно-экономический колледж им. Н. П. Пастухова\"\n" +
                "150046, г. Ярославль, ул. Гагарина, 8\n" +
                "44-26-77\n\n" +
                "« 11 » июня 2025 г   № 6"));
            header.FontFamily = new FontFamily("Times New Roman");
            header.FontSize = 10;
            header.Margin = new Thickness(0, 0, 0, 20);
            document.Blocks.Add(header);

            // Заголовок "СПРАВКА"
            Paragraph title = new Paragraph();
            title.Inlines.Add(new Run("СПРАВКА"));
            title.FontFamily = new FontFamily("Times New Roman");
            title.FontSize = 14;
            title.FontWeight = FontWeights.Bold;
            title.TextAlignment = TextAlignment.Center;
            title.Margin = new Thickness(0, 0, 0, 20);
            document.Blocks.Add(title);

            // Основной текст
            Paragraph content = new Paragraph();
            content.Inlines.Add(new Run(
                "Выдана настоящая Комарову Ивану Сергеевичу в том, что он(а) обучается " +
                "в ГПОАУ ЯО «Ярославский промышленно-экономический колледж им. Н. П. Пастухова» " +
                "на 2 курсе по очной форме обучения, на бюджетной основе.\n\n" +
                "Выдана для предъявления по месту требования.\n\n" +
                "Срок обучения с 01.09.2021 по 30.06.2025."));
            content.FontFamily = new FontFamily("Times New Roman");
            content.FontSize = 11;
            content.TextAlignment = TextAlignment.Left;
            content.Margin = new Thickness(0, 0, 0, 40);
            document.Blocks.Add(content);

            // Подпись
            Table signTable = new Table();
            signTable.Margin = new Thickness(0, 0, 0, 0);
            signTable.TextAlignment = TextAlignment.Center;
            signTable.Columns.Add(new TableColumn { Width = new GridLength(250) });
            signTable.Columns.Add(new TableColumn { Width = new GridLength(250) });

            TableRowGroup group = new TableRowGroup();
            TableRow row = new TableRow();

            // Левая часть подписи
            TableCell cellLeft = new TableCell();
            cellLeft.Blocks.Add(new Paragraph(new Run("Заведующий отделением")));
            cellLeft.TextAlignment = TextAlignment.Left;
            row.Cells.Add(cellLeft);

            // Правая часть подписи
            TableCell cellRight = new TableCell();
            cellRight.Blocks.Add(new Paragraph(new Run("Ю. В. Маянцева")));
            cellRight.TextAlignment = TextAlignment.Right;
            row.Cells.Add(cellRight);

            group.Rows.Add(row);
            signTable.RowGroups.Add(group);
            document.Blocks.Add(signTable);

            // Телефон
            Paragraph phone = new Paragraph();
            phone.Inlines.Add(new Run("Тел. 48-05-24"));
            phone.FontFamily = new FontFamily("Times New Roman");
            phone.FontSize = 11;
            phone.Margin = new Thickness(0, 20, 0, 0);
            document.Blocks.Add(phone);

            return document;
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