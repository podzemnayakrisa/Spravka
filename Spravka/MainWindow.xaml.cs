using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
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
        private const string GoogleScriptUrl = "https://script.google.com/macros/s/AKfycbzLfhuXtJycyH67GGTITeJGuxxoNSOXXBT5U2GjLsQAcNmvDKUbOjF5RwaYZdnN7qJ3gg/exec";

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
                await UpdateStatusInGoogleSheet(
                    selectedItem.Email,
                    selectedItem.Status,
                    selectedItem.FullName // Важно передать полное имя
                );
            }
        }

        private async Task UpdateStatusInGoogleSheet(string email, string status, string fullName)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var data = new
                    {
                        action = "update_status",
                        email,
                        status,
                        fullName // Убедитесь, что передается
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
                MessageBox.Show($"Ошибка при обновлении статуса: {ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
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
            if (ResponsesDataGrid.SelectedItem is ResponseItem selectedItem)
            {
                try
                {
                    var result = await CreateCertificateInGoogle(selectedItem);

                    if (result.Success)
                    {
                        selectedItem.PdfUrl = result.PdfUrl;
                        ResponsesDataGrid.Items.Refresh();
                        MessageBox.Show("Справка успешно создана!");
                    }
                    else
                    {
                        MessageBox.Show($"Ошибка при создании справки: {result.Error}");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при создании справки: {ex.Message}");
                }
            }
        }

        private async void PrintCertificateMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (ResponsesDataGrid.SelectedItem is ResponseItem selectedItem)
            {
                if (string.IsNullOrEmpty(selectedItem.PdfUrl))
                {
                    MessageBox.Show("Для выбранной записи отсутствует PDF-файл. Создайте справку сначала.");
                    return;
                }

                try
                {
                    string tempFilePath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.pdf");

                    using (var client = new WebClient())
                    {
                        await client.DownloadFileTaskAsync(selectedItem.PdfUrl, tempFilePath);
                    }

                    PrintUsingSystemDialog(tempFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при печати: {ex.Message}");
                }
            }
        }

        private void PrintUsingSystemDialog(string filePath)
        {
            try
            {
                // Проверка существования файла
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Файл для печати не найден");
                    return;
                }

                // Открываем файл в программе по умолчанию с параметром печати
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    Verb = "print",
                    UseShellExecute = true,
                    CreateNoWindow = true
                });

                // Отложенное удаление файла
                Task.Delay(30000).ContinueWith(t =>
                {
                    try
                    {
                        if (File.Exists(filePath))
                            File.Delete(filePath);
                    }
                    catch { /* Игнорируем ошибки удаления */ }
                });
            }
            catch (Exception ex)
            {
                // Попробуем альтернативный метод печати
                TryAlternativePrint(filePath, ex);
            }
        }

        private void TryAlternativePrint(string filePath, Exception originalEx)
        {
            try
            {
                // Попытка 1: Открыть в ассоциированной программе
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });

                MessageBox.Show("Файл открыт для просмотра. Пожалуйста, распечатайте его вручную.");
            }
            catch
            {
                // Попытка 2: Использовать системную утилиту Microsoft Print to PDF
                PrintWithBuiltInFeature(filePath, originalEx);
            }
        }

        private void PrintWithBuiltInFeature(string filePath, Exception originalEx)
        {
            try
            {
                string printerName = "Microsoft Print to PDF";

                // Проверяем наличие принтера
                if (!PrinterSettings.InstalledPrinters.Cast<string>().Any(p => p.Contains(printerName)))
                {
                    throw new Exception($"Принтер '{printerName}' не найден");
                }

                using (var document = PdfiumViewer.PdfDocument.Load(filePath))
                {
                    using (var printDocument = document.CreatePrintDocument())
                    {
                        printDocument.PrinterSettings.PrinterName = printerName;
                        printDocument.Print();
                    }
                }
            }
            catch (Exception ex)
            {
                // Если все методы не сработали, показываем оригинальную ошибку
                MessageBox.Show($"Ошибка печати: {originalEx.Message}\n\nДополнительно: {ex.Message}");
            }
        }
        private async Task CreateCertificateForStudent(ResponseItem student, bool showMessage = true)
        {
            try
            {
                var result = await CreateCertificateInGoogle(student);

                if (result.Success)
                {
                    // Автоматическое сохранение без диалога
                    string fileName = $"Справка_{student.FullName}_{DateTime.Now:yyyyMMddHHmmss}.pdf";
                    string savePath = Path.Combine(GetSaveFolder(), fileName);
                    await DownloadPdf(result.PdfUrl, savePath);

                    // Открываем предпросмотр
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = savePath,
                        UseShellExecute = true
                    });

                    if (showMessage)
                    {
                        MessageBox.Show($"Справка сохранена: {savePath}", "Успех",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                else
                {
                    MessageBox.Show($"Ошибка при создании справки: {result.Error}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string GetSaveFolder()
        {
            // Указываем нужный путь
            string path = @"D:\пдф\Spravka";

            // Создаем папку если не существует
            if (!Directory.Exists(path))
            {
                try
                {
                    Directory.CreateDirectory(path);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось создать папку: {ex.Message}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);

                    // Возвращаем временную папку в случае ошибки
                    return Path.GetTempPath();
                }
            }

            return path;
        }

        private async Task DownloadPdf(string pdfUrl, string savePath)
        {
            using (var client = new HttpClient())
            {
                var response = await client.GetAsync(pdfUrl);
                response.EnsureSuccessStatusCode();

                using (var fs = new FileStream(savePath, FileMode.Create))
                {
                    await response.Content.CopyToAsync(fs);
                }
            }
        }

        private async Task PrintCertificateForStudent(ResponseItem student)
        {
            try
            {
                var result = await CreateCertificateInGoogle(student);

                if (result.Success)
                {
                    // Автоматическое сохранение во временный файл
                    string tempFile = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid()}.pdf");
                    await DownloadPdf(result.PdfUrl, tempFile);

                    // Печать без диалога
                    PrintUsingSystemDialog(tempFile);
                }
                else
                {
                    MessageBox.Show($"Ошибка: {result.Error}", "Ошибка",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка печати: {ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void PreviewMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (ResponsesDataGrid.SelectedItem is ResponseItem student)
            {
                CreateCertificateForStudent(student, showMessage: false).ConfigureAwait(false);
            }
        }

        private void PrintMenuItem_Click(object sender, RoutedEventArgs e)
        {
            if (ResponsesDataGrid.SelectedItem is ResponseItem student)
            {
                PrintCertificateForStudent(student).ConfigureAwait(false);
            }
        }

        private async Task<CertificateResponse> CreateCertificateInGoogle(ResponseItem student)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var data = new
                    {
                        action = "create_certificate",
                        fullName = student.FullName,
                        course = student.Course,
                        educationForm = student.EducationForm,
                        basis = student.Basis,
                        group = student.Group
                    };

                    var content = new StringContent(JsonConvert.SerializeObject(data),
                        Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(GoogleScriptUrl, content);

                    if (!response.IsSuccessStatusCode)
                    {
                        return new CertificateResponse
                        {
                            Success = false,
                            Error = $"HTTP ошибка: {response.StatusCode}"
                        };
                    }

                    string json = await response.Content.ReadAsStringAsync();
                    var result = JsonConvert.DeserializeObject<GoogleScriptResponse<CertificateResponse>>(json);

                    // Проверяем структуру ответа
                    if (result.Success && result.Data != null)
                    {
                        return new CertificateResponse
                        {
                            Success = true,
                            DocumentUrl = result.Data.DocumentUrl,
                            PdfUrl = result.Data.PdfUrl,
                            CertificateNumber = result.Data.CertificateNumber,
                            Error = result.Error
                        };
                    }

                    return new CertificateResponse
                    {
                        Success = result.Success,
                        Error = result.Error ?? "Неизвестная ошибка при создании сертификата"
                    };
                }
            }
            catch (Exception ex)
            {
                return new CertificateResponse
                {
                    Success = false,
                    Error = ex.Message
                };
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

    public class CertificateResponse
    {
        public bool Success { get; set; }
        public string DocumentUrl { get; set; }
        public string PdfUrl { get; set; }
        public string CertificateNumber { get; set; }
        public string Error { get; set; }
    }

    public class GoogleScriptResponse<T>
    {
        public bool Success { get; set; }
        public T Data { get; set; }
        public string Error { get; set; }
    }
}